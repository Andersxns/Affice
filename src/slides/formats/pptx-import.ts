/**
 * PowerPoint (.pptx) reader. Resolves the full inheritance chain (presentation defaults → master text
 * styles → master/layout placeholders → shape → paragraph → run) into explicit properties, keeps theme
 * colours and fonts as references, and reads pictures, tables, charts, groups, SmartArt drawings,
 * notes, transitions and animations.
 */
import { strFromU8, unzipSync, type Unzipped } from 'fflate';
import { bytesToDataUrl, mimeFromExt } from '@/lib/utils';
import {
  hslToRgb,
  lumModOff,
  newId,
  rgbToHsl,
  type Anim,
  type AnimClass,
  type AnimEffect,
  type ArrowType,
  type Bullet,
  type CustomPath,
  type Dash,
  type Dir,
  type El,
  type Fill,
  type GroupEl,
  type ImageEl,
  type Layout,
  type LayoutType,
  type Line,
  type NumStyle,
  type Para,
  type PathCmd,
  type PlaceholderType,
  type Presentation,
  type Run,
  type Shadow,
  type ShapeEl,
  type Slide,
  type TableCell,
  type TableEl,
  type TextBody,
  type Theme,
  type ThemeColorName,
  type Transition,
  type TransitionType,
} from '../model';
import { styleFromId } from '../tables';
import type { LoadedPresentation } from './index';
import { parseChartXml } from './pptx-charts';

const EMU = 9525;
const px = (emu: string | null | undefined) => (emu ? Number(emu) / EMU : 0);
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/* ================================================================ XML */

function kids(el: Element | null | undefined, name?: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  for (const c of Array.from(el.children)) if (!name || c.localName === name) out.push(c);
  return out;
}
function kid(el: Element | null | undefined, name: string): Element | null {
  if (!el) return null;
  for (const c of Array.from(el.children)) if (c.localName === name) return c;
  return null;
}
function path(el: Element | null | undefined, ...names: string[]): Element | null {
  let cur: Element | null | undefined = el;
  for (const n of names) {
    cur = kid(cur, n);
    if (!cur) return null;
  }
  return cur ?? null;
}
const attr = (el: Element | null | undefined, name: string) => el?.getAttribute(name) ?? null;
const num = (el: Element | null | undefined, name: string, def = 0) => {
  const v = attr(el, name);
  return v === null ? def : Number(v);
};
const rid = (el: Element | null | undefined, name = 'embed') => el?.getAttributeNS(NS_R, name) ?? el?.getAttribute(`r:${name}`) ?? null;

/** Picks the understood branch of mc:AlternateContent. */
function unwrap(el: Element): Element[] {
  if (el.localName !== 'AlternateContent') return [el];
  const choice = kids(el, 'Choice').find((c) => /^(p14|p15|a14|p159|asvg)$/.test(attr(c, 'Requires') ?? '') || !attr(c, 'Requires'));
  const branch = choice && kids(choice).some((k) => ['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp', 'transition'].includes(k.localName)) ? choice : kid(el, 'Fallback');
  return kids(branch ?? el);
}

/* ============================================================== package */

class Pkg {
  constructor(public zip: Unzipped) {}
  has(p: string): boolean {
    return !!this.zip[p];
  }
  text(p: string): string | null {
    const d = this.zip[p];
    return d ? strFromU8(d) : null;
  }
  xml(p: string): Document | null {
    const t = this.text(p);
    return t ? new DOMParser().parseFromString(t, 'application/xml') : null;
  }
  rels(p: string): Map<string, { target: string; type: string; external: boolean }> {
    const dir = p.split('/').slice(0, -1).join('/');
    const file = `${dir}/_rels/${p.split('/').pop()}.rels`;
    const out = new Map<string, { target: string; type: string; external: boolean }>();
    const doc = this.xml(file);
    if (!doc) return out;
    for (const r of Array.from(doc.getElementsByTagName('Relationship'))) {
      const id = r.getAttribute('Id');
      const target = r.getAttribute('Target') ?? '';
      const external = r.getAttribute('TargetMode') === 'External';
      if (!id) continue;
      out.set(id, { target: external ? target : resolve(p, target), type: (r.getAttribute('Type') ?? '').split('/').pop() ?? '', external });
    }
    return out;
  }
}

function resolve(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

/* ============================================================== colours */

const PRESET: Record<string, string> = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', gray: '#808080', grey: '#808080', ltGray: '#D3D3D3', dkGray: '#A9A9A9', orange: '#FFA500', purple: '#800080', cyan: '#00FFFF', magenta: '#FF00FF', navy: '#000080', maroon: '#800000' };

interface ColorEnv {
  scheme: Record<string, string>;
  /** Slide colour map (bg1 → lt1 …). */
  map: Record<string, string>;
  /** Colour substituted for phClr (style references). */
  ph?: string;
}

const MY_NAMES = new Set(['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']);

function hex2(n: number): string {
  return Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0');
}

function applyMods(hex: string, el: Element): string {
  let out = hex;
  for (const m of kids(el)) {
    const v = num(m, 'val') / 100000;
    switch (m.localName) {
      case 'lumMod': {
        const off = kid(el, 'lumOff');
        out = lumModOff(out, v, off ? num(off, 'val') / 100000 : 0);
        break;
      }
      case 'tint': {
        const [r, g, b] = [1, 3, 5].map((i) => parseInt(out.slice(i, i + 2), 16));
        out = `#${hex2(r + (255 - r) * (1 - v))}${hex2(g + (255 - g) * (1 - v))}${hex2(b + (255 - b) * (1 - v))}`;
        break;
      }
      case 'shade': {
        const [r, g, b] = [1, 3, 5].map((i) => parseInt(out.slice(i, i + 2), 16));
        out = `#${hex2(r * v)}${hex2(g * v)}${hex2(b * v)}`;
        break;
      }
      case 'satMod': {
        const [r, g, b] = [1, 3, 5].map((i) => parseInt(out.slice(i, i + 2), 16));
        const [h, s, l] = rgbToHsl(r, g, b);
        const [nr, ng, nb] = hslToRgb(h, Math.min(1, s * v), l);
        out = `#${hex2(nr)}${hex2(ng)}${hex2(nb)}`;
        break;
      }
    }
  }
  return out.toUpperCase();
}

/** Theme-aware colour value from a DrawingML colour element (schemeClr, srgbClr…). */
function colorOf(el: Element | null | undefined, env: ColorEnv): string | undefined {
  if (!el) return undefined;
  const c = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'].includes(el.localName) ? el : kids(el).find((k) => /Clr$/.test(k.localName));
  if (!c) return undefined;
  const alphaEl = kid(c, 'alpha');
  const alpha = alphaEl ? num(alphaEl, 'val') / 100000 : 1;
  const withAlpha = (v: string) => (alpha < 1 ? (v.startsWith('@') ? `${v}/${Math.round(alpha * 100)}` : `${v}${hex2(alpha * 255)}`) : v);
  const mods = kids(c).filter((m) => m.localName !== 'alpha');
  if (c.localName === 'schemeClr') {
    let name = attr(c, 'val') ?? 'tx1';
    if (name === 'phClr') {
      if (!env.ph) return undefined;
      if (!mods.length) return env.ph.startsWith('@') || alpha >= 1 ? env.ph : withAlpha(env.ph);
      // modifiers on the placeholder colour: resolve
      const base = env.ph.startsWith('@') ? resolveRef(env.ph, env) : env.ph.slice(0, 7);
      return withAlpha(applyMods(base, c));
    }
    // dk1/lt1… map back to the slide's roles through the colour map
    if (!MY_NAMES.has(name)) {
      const role = Object.entries(env.map).find(([, v]) => v === name)?.[0];
      name = role && MY_NAMES.has(role) ? role : name === 'dk1' ? 'tx1' : name === 'lt1' ? 'bg1' : name === 'dk2' ? 'tx2' : name === 'lt2' ? 'bg2' : name;
    }
    const lumMod = kid(c, 'lumMod');
    const lumOff = kid(c, 'lumOff');
    const others = mods.filter((m) => m.localName !== 'lumMod' && m.localName !== 'lumOff');
    if (!others.length) {
      const lm = lumMod ? num(lumMod, 'val') / 1000 : 100;
      const lo = lumOff ? num(lumOff, 'val') / 1000 : 0;
      if (!lumMod && !lumOff) return withAlpha(`@${name}`);
      if (lo > 0 && Math.abs(lm + lo - 100) < 1.5) return withAlpha(`@${name}+${Math.round(lo)}`);
      if (lo === 0 && lm < 100) return withAlpha(`@${name}-${Math.round(100 - lm)}`);
    }
    return withAlpha(applyMods(resolveRef(`@${name}`, env), c));
  }
  let base = '#000000';
  if (c.localName === 'srgbClr') base = `#${(attr(c, 'val') ?? '000000').toUpperCase()}`;
  else if (c.localName === 'sysClr') base = `#${(attr(c, 'lastClr') ?? (attr(c, 'val') === 'window' ? 'FFFFFF' : '000000')).toUpperCase()}`;
  else if (c.localName === 'prstClr') base = PRESET[attr(c, 'val') ?? 'black'] ?? '#000000';
  else if (c.localName === 'scrgbClr') base = `#${hex2((num(c, 'r') / 100000) * 255)}${hex2((num(c, 'g') / 100000) * 255)}${hex2((num(c, 'b') / 100000) * 255)}`;
  else if (c.localName === 'hslClr') {
    const [r, g, b] = hslToRgb(num(c, 'hue') / 21600000, num(c, 'sat') / 100000, num(c, 'lum') / 100000);
    base = `#${hex2(r)}${hex2(g)}${hex2(b)}`;
  }
  return withAlpha(mods.length ? applyMods(base, c) : base);
}

function resolveRef(ref: string, env: ColorEnv): string {
  const m = /^@(\w+)([+-]\d+)?/.exec(ref);
  if (!m) return '#000000';
  const name = m[1];
  const slot = MY_NAMES.has(name) ? (env.map[name] ?? name) : name;
  const base = env.scheme[slot] ?? env.scheme[name] ?? '#000000';
  const lum = m[2] ? parseInt(m[2], 10) : 0;
  return lum > 0 ? lumModOff(base, 1 - lum / 100, lum / 100) : lum < 0 ? lumModOff(base, 1 + lum / 100, 0) : base;
}

/* ================================================================ theme */

interface ThemeInfo {
  name: string;
  scheme: Record<string, string>;
  major: string;
  minor: string;
  fills: Element[];
  lines: Element[];
  bgFills: Element[];
}

function readTheme(doc: Document | null): ThemeInfo {
  const scheme: Record<string, string> = { dk1: '#000000', lt1: '#FFFFFF', dk2: '#44546A', lt2: '#E7E6E6', accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47', hlink: '#0563C1', folHlink: '#954F72' };
  let major = 'Calibri Light';
  let minor = 'Calibri';
  let name = 'Imported theme';
  const fills: Element[] = [];
  const lines: Element[] = [];
  const bgFills: Element[] = [];
  const root = doc?.documentElement;
  if (root) {
    name = attr(root, 'name') ?? name;
    const cs = path(root, 'themeElements', 'clrScheme');
    for (const c of kids(cs)) {
      const env: ColorEnv = { scheme: {}, map: {} };
      const v = colorOf(c, env);
      if (v && !v.startsWith('@')) scheme[c.localName] = v.slice(0, 7);
    }
    const fs = path(root, 'themeElements', 'fontScheme');
    major = attr(path(fs, 'majorFont', 'latin'), 'typeface') || major;
    minor = attr(path(fs, 'minorFont', 'latin'), 'typeface') || minor;
    const fmt = path(root, 'themeElements', 'fmtScheme');
    fills.push(...kids(kid(fmt, 'fillStyleLst')));
    lines.push(...kids(kid(fmt, 'lnStyleLst')));
    bgFills.push(...kids(kid(fmt, 'bgFillStyleLst')));
  }
  return { name, scheme, major, minor, fills, lines, bgFills };
}

/* ============================================================ contexts */

interface LevelProps {
  p: Record<string, unknown>;
  r: Omit<Run, 'text'>;
}

type ListStyle = LevelProps[];

interface PhInfo {
  type: string;
  idx: string;
  xfrm: Element | null;
  body: Element | null;
  lst: ListStyle;
  spPr: Element | null;
}

interface MasterInfo {
  path: string;
  env: ColorEnv;
  title: ListStyle;
  body: ListStyle;
  other: ListStyle;
  phs: PhInfo[];
  bg: Fill | undefined;
  decor: El[];
}

interface LayoutInfo {
  path: string;
  layout: Layout;
  phs: PhInfo[];
  master: MasterInfo;
  showMasterSp: boolean;
}

interface Ctx {
  pkg: Pkg;
  theme: ThemeInfo;
  env: ColorEnv;
  partPath: string;
  rels: Map<string, { target: string; type: string; external: boolean }>;
  slides: string[];
  media: Map<string, string>;
  warnings: Set<string>;
}

/* ================================================================ fills */

function mediaUrl(ctx: Ctx, r: string | null): string | null {
  if (!r) return null;
  const rel = ctx.rels.get(r);
  if (!rel || rel.external) return null;
  const hit = ctx.media.get(rel.target);
  if (hit) return hit;
  const data = ctx.pkg.zip[rel.target];
  if (!data) return null;
  const ext = rel.target.split('.').pop()?.toLowerCase() ?? 'png';
  if (ext === 'emf' || ext === 'wmf') ctx.warnings.add('Some pictures use Windows metafile formats (EMF/WMF) that can only be shown as placeholders.');
  const url = bytesToDataUrl(data, mimeFromExt(ext === 'jpg' ? 'jpeg' : ext));
  ctx.media.set(rel.target, url);
  return url;
}

function fillOf(el: Element | null | undefined, ctx: Ctx, env: ColorEnv): Fill | undefined {
  if (!el) return undefined;
  const f = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(el.localName) ? el : kids(el).find((k) => ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(k.localName));
  if (!f) return undefined;
  switch (f.localName) {
    case 'noFill':
      return { type: 'none' };
    case 'solidFill': {
      const c = colorOf(f, env);
      return c ? { type: 'solid', color: c } : undefined;
    }
    case 'gradFill': {
      const stops = kids(kid(f, 'gsLst'), 'gs').map((g) => ({ pos: num(g, 'pos') / 100000, color: colorOf(g, env) ?? '#000000' }));
      if (!stops.length) return undefined;
      const lin = kid(f, 'lin');
      const pathEl = kid(f, 'path');
      return { type: 'gradient', stops, angle: lin ? num(lin, 'ang') / 60000 : 90, radial: !!pathEl && !lin };
    }
    case 'blipFill': {
      const url = mediaUrl(ctx, rid(kid(f, 'blip')));
      return url ? { type: 'image', src: url, mode: kid(f, 'tile') ? 'tile' : 'stretch' } : undefined;
    }
    case 'pattFill': {
      const c = colorOf(kid(f, 'fgClr'), env);
      return c ? { type: 'solid', color: c } : undefined;
    }
    default:
      return undefined;
  }
}

const DASHES: Record<string, Dash> = { dash: 'dash', sysDash: 'sysDash', dot: 'dot', sysDot: 'sysDot', dashDot: 'dashDot', sysDashDot: 'dashDot', lgDash: 'longDash', lgDashDot: 'dashDot', lgDashDotDot: 'dashDot', sysDashDotDot: 'dashDot' };

function lineOf(ln: Element | null, ctx: Ctx, env: ColorEnv, base?: Line | null): Line | null | undefined {
  if (!ln) return base;
  if (kid(ln, 'noFill')) return null;
  const f = fillOf(ln, ctx, env);
  const color = f?.type === 'solid' ? f.color : f?.type === 'gradient' ? f.stops[0].color : (base?.color ?? undefined);
  const wAttr = attr(ln, 'w');
  const width = wAttr !== null ? Number(wAttr) / EMU : (base?.width ?? 0.75);
  if (!color) return base;
  const out: Line = { color, width: Math.max(0.5, width) };
  const dash = attr(kid(ln, 'prstDash'), 'val');
  if (dash && dash !== 'solid') out.dash = DASHES[dash] ?? 'dash';
  const arrow = (e: Element | null): ArrowType | undefined => {
    const t = attr(e, 'type');
    return t && t !== 'none' ? ((['triangle', 'arrow', 'stealth', 'oval', 'diamond'] as ArrowType[]).includes(t as ArrowType) ? (t as ArrowType) : 'triangle') : undefined;
  };
  const head = arrow(kid(ln, 'headEnd'));
  const tail = arrow(kid(ln, 'tailEnd'));
  if (head) out.head = head;
  if (tail) out.tail = tail;
  const cap = attr(ln, 'cap');
  if (cap === 'rnd') out.cap = 'round';
  else if (cap === 'sq') out.cap = 'square';
  return out;
}

function shadowOf(spPr: Element | null, env: ColorEnv): Shadow | undefined {
  const sh = path(spPr, 'effectLst', 'outerShdw');
  if (!sh) return undefined;
  return { color: colorOf(sh, env) ?? '#00000066', blur: px(attr(sh, 'blurRad')), dist: px(attr(sh, 'dist')), angle: num(sh, 'dir') / 60000 };
}

/* ================================================================= text */

const ALIGN: Record<string, Para['align']> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', justLow: 'justify' };

function levelProps(lvl: Element | null, ctx: Ctx, env: ColorEnv): LevelProps {
  const p: Record<string, unknown> = {};
  const r: Omit<Run, 'text'> = {};
  if (!lvl) return { p, r };
  const a = attr(lvl, 'algn');
  if (a) p.align = ALIGN[a] ?? 'left';
  if (attr(lvl, 'marL') !== null) p.marL = px(attr(lvl, 'marL'));
  if (attr(lvl, 'indent') !== null) p.indent = px(attr(lvl, 'indent'));
  const ln = kid(lvl, 'lnSpc');
  if (ln) {
    const pct = kid(ln, 'spcPct');
    const pts = kid(ln, 'spcPts');
    if (pct) p.lineSpacing = num(pct, 'val') / 100000;
    else if (pts) p.lineSpacingPts = num(pts, 'val') / 100;
  }
  const sp = (name: string, key: string) => {
    const e = kid(lvl, name);
    if (!e) return;
    const pts = kid(e, 'spcPts');
    const pct = kid(e, 'spcPct');
    if (pts) p[key] = num(pts, 'val') / 100;
    else if (pct) p[`${key}Pct`] = num(pct, 'val') / 100000;
  };
  sp('spcBef', 'spaceBefore');
  sp('spcAft', 'spaceAfter');
  if (kid(lvl, 'buNone')) p.bullet = { type: 'none' };
  const buChar = kid(lvl, 'buChar');
  if (buChar) p.bullet = { type: 'char', char: attr(buChar, 'char') ?? '•' };
  const buNum = kid(lvl, 'buAutoNum');
  if (buNum) p.bullet = { type: 'num', style: numStyle(attr(buNum, 'type')), start: num(buNum, 'startAt', 1) };
  if (kid(lvl, 'buBlip')) p.bullet = { type: 'char', char: '•' };
  const buClr = kid(lvl, 'buClr');
  if (buClr) p.buColor = colorOf(buClr, env);
  const buFont = attr(kid(lvl, 'buFont'), 'typeface');
  if (buFont) p.buFont = buFont;
  const def = kid(lvl, 'defRPr');
  if (def) Object.assign(r, runProps(def, ctx, env));
  return { p, r };
}

function numStyle(t: string | null): NumStyle {
  switch (t) {
    case 'arabicParenR':
    case 'arabicParenBoth':
      return 'arabicParenR';
    case 'alphaLcPeriod':
      return 'alphaLcPeriod';
    case 'alphaUcPeriod':
      return 'alphaUcPeriod';
    case 'alphaLcParenR':
    case 'alphaLcParenBoth':
    case 'alphaUcParenR':
      return 'alphaLcParenR';
    case 'romanLcPeriod':
      return 'romanLcPeriod';
    case 'romanUcPeriod':
      return 'romanUcPeriod';
    default:
      return 'arabicPeriod';
  }
}

function fontRef(tf: string | null): string | undefined {
  if (!tf) return undefined;
  if (tf === '+mj-lt' || tf === '+mj-ea' || tf === '+mj-cs') return '+major';
  if (tf === '+mn-lt' || tf === '+mn-ea' || tf === '+mn-cs') return '+minor';
  return tf;
}

function runProps(rPr: Element, ctx: Ctx, env: ColorEnv): Omit<Run, 'text'> {
  const r: Omit<Run, 'text'> = {};
  const sz = attr(rPr, 'sz');
  if (sz) r.size = Number(sz) / 100;
  const b = attr(rPr, 'b');
  if (b !== null) r.b = b === '1' || b === 'true';
  const i = attr(rPr, 'i');
  if (i !== null) r.i = i === '1' || i === 'true';
  const u = attr(rPr, 'u');
  if (u !== null) r.u = u !== 'none';
  const st = attr(rPr, 'strike');
  if (st !== null) r.s = st !== 'noStrike';
  const bl = attr(rPr, 'baseline');
  if (bl !== null) {
    const v = Number(bl);
    r.sup = v > 0;
    r.sub = v < 0;
  }
  const cap = attr(rPr, 'cap');
  if (cap !== null) r.caps = cap === 'all';
  const spc = attr(rPr, 'spc');
  if (spc) r.spacing = Number(spc) / 100;
  const fill = fillOf(kids(rPr).find((k) => ['solidFill', 'gradFill', 'noFill'].includes(k.localName)) ?? null, ctx, env);
  if (fill?.type === 'solid') r.color = fill.color;
  else if (fill?.type === 'gradient') r.color = fill.stops[0].color;
  const hl = kid(rPr, 'highlight');
  if (hl) r.hl = colorOf(hl, env);
  const latin = fontRef(attr(kid(rPr, 'latin'), 'typeface'));
  if (latin) r.font = latin;
  const link = kid(rPr, 'hlinkClick');
  if (link) {
    const l = linkOf(link, ctx);
    if (l) r.link = l;
  }
  return r;
}

function linkOf(el: Element, ctx: Ctx): string | undefined {
  const action = attr(el, 'action') ?? '';
  const jump = /jump=(\w+)/.exec(action)?.[1];
  if (jump) return { nextslide: '#next', previousslide: '#prev', firstslide: '#first', lastslide: '#last' }[jump];
  const r = rid(el, 'id');
  const rel = r ? ctx.rels.get(r) : undefined;
  if (!rel) return undefined;
  if (rel.external) return rel.target;
  const idx = ctx.slides.indexOf(rel.target);
  return idx >= 0 ? `#slide:${idx + 1}` : undefined;
}

function listStyle(lst: Element | null, ctx: Ctx, env: ColorEnv): ListStyle {
  const out: ListStyle = [];
  for (let i = 0; i < 9; i++) out.push(levelProps(kid(lst, `lvl${i + 1}pPr`), ctx, env));
  return out;
}

function mergeLevels(...lists: (ListStyle | undefined)[]): ListStyle {
  const out: ListStyle = [];
  for (let i = 0; i < 9; i++) {
    const p: Record<string, unknown> = {};
    const r: Omit<Run, 'text'> = {};
    for (const l of lists) {
      if (!l) continue;
      Object.assign(p, l[i].p);
      Object.assign(r, l[i].r);
    }
    out.push({ p, r });
  }
  return out;
}

interface BodyProps {
  anchor?: TextBody['anchor'];
  inset?: [number, number, number, number];
  autofit?: TextBody['autofit'];
  fontScale?: number;
  wrap?: boolean;
  vert?: TextBody['vert'];
  columns?: number;
}

function bodyProps(bp: Element | null): BodyProps {
  const out: BodyProps = {};
  if (!bp) return out;
  const anchor = attr(bp, 'anchor');
  if (anchor) out.anchor = anchor === 'ctr' ? 'm' : anchor === 'b' ? 'b' : 't';
  const ins = ['lIns', 'tIns', 'rIns', 'bIns'].map((k) => attr(bp, k));
  if (ins.some((v) => v !== null)) {
    const d = [9.6, 4.8, 9.6, 4.8];
    out.inset = ins.map((v, i) => (v === null ? d[i] : px(v))) as [number, number, number, number];
  }
  if (kid(bp, 'normAutofit')) {
    out.autofit = 'shrink';
    const fs = attr(kid(bp, 'normAutofit'), 'fontScale');
    if (fs) out.fontScale = Number(fs) / 100000;
  } else if (kid(bp, 'spAutoFit')) out.autofit = 'resize';
  else if (kid(bp, 'noAutofit')) out.autofit = 'none';
  const wrap = attr(bp, 'wrap');
  if (wrap) out.wrap = wrap !== 'none';
  const vert = attr(bp, 'vert');
  if (vert === 'vert' || vert === 'eaVert') out.vert = 'vert';
  else if (vert === 'vert270') out.vert = 'vert270';
  const cols = attr(bp, 'numCol');
  if (cols && Number(cols) > 1) out.columns = Number(cols);
  return out;
}

/** Converts a txBody using the inherited level styles into an explicit text body. */
function textOf(txBody: Element | null, inherited: ListStyle, inheritedBody: BodyProps, ctx: Ctx, env: ColorEnv, baseRun: Omit<Run, 'text'> = {}): TextBody | undefined {
  if (!txBody) return undefined;
  const own = listStyle(kid(txBody, 'lstStyle'), ctx, env);
  const levels = mergeLevels(inherited, own);
  const bp = { ...inheritedBody, ...bodyProps(kid(txBody, 'bodyPr')) };
  const paras: Para[] = [];
  for (const p of kids(txBody, 'p')) {
    const pPr = kid(p, 'pPr');
    const lvl = Math.max(0, Math.min(8, num(pPr, 'lvl', 0)));
    const lp = levels[lvl];
    const pp = { ...lp.p, ...levelProps(pPr, ctx, env).p };
    const defRun: Omit<Run, 'text'> = { ...baseRun, ...lp.r, ...(pPr && kid(pPr, 'defRPr') ? runProps(kid(pPr, 'defRPr')!, ctx, env) : {}) };
    const runs: Run[] = [];
    for (const r of kids(p)) {
      if (r.localName === 'r' || r.localName === 'fld') {
        const rp = kid(r, 'rPr');
        const props = { ...defRun, ...(rp ? runProps(rp, ctx, env) : {}) };
        const text = kid(r, 't')?.textContent ?? '';
        const run: Run = { ...props, text };
        if (r.localName === 'fld') {
          const t = attr(r, 'type') ?? '';
          if (t === 'slidenum') run.field = 'slidenum';
          else if (t.startsWith('datetime')) run.field = 'date';
        }
        runs.push(run);
      } else if (r.localName === 'br') {
        if (runs.length) runs[runs.length - 1].text += '\n';
        else runs.push({ ...defRun, text: '\n' });
      }
    }
    const endRp = kid(p, 'endParaRPr');
    const endRun: Omit<Run, 'text'> = { ...defRun, ...(endRp ? runProps(endRp, ctx, env) : {}) };
    const size = runs[0]?.size ?? endRun.size ?? 18;
    const para: Para = {
      runs: runs.map((r) => cleanRun(r) as Run),
      level: lvl || undefined,
      align: (pp.align as Para['align']) ?? 'left',
      marL: (pp.marL as number) ?? 0,
      indent: (pp.indent as number) ?? 0,
      lineSpacing: pp.lineSpacing !== undefined ? (pp.lineSpacing as number) : pp.lineSpacingPts !== undefined ? Math.max(0.5, (pp.lineSpacingPts as number) / (size * 1.2)) : 1,
      spaceBefore: pp.spaceBefore !== undefined ? (pp.spaceBefore as number) : pp.spaceBeforePct !== undefined ? (pp.spaceBeforePct as number) * size : 0,
      spaceAfter: pp.spaceAfter !== undefined ? (pp.spaceAfter as number) : pp.spaceAfterPct !== undefined ? (pp.spaceAfterPct as number) * size : 0,
      bullet: bulletOf(pp),
    };
    if (!runs.length || !runs.some((r) => r.text)) para.endRun = cleanRun(endRun);
    paras.push(para);
  }
  const body: TextBody = { paras: paras.length ? paras : [{ runs: [] }] };
  if (bp.anchor) body.anchor = bp.anchor;
  if (bp.inset) body.inset = bp.inset;
  body.autofit = bp.autofit ?? 'none';
  if (bp.fontScale && bp.fontScale < 1) body.fontScale = bp.fontScale;
  if (bp.wrap === false) body.wrap = false;
  if (bp.vert) body.vert = bp.vert;
  if (bp.columns) body.columns = bp.columns;
  return body;
}

function bulletOf(pp: Record<string, unknown>): Bullet {
  const b = pp.bullet as Bullet | undefined;
  if (!b || b.type === 'none') return { type: 'none' };
  if (b.type === 'char') return { ...b, ...(pp.buColor ? { color: pp.buColor as string } : {}), ...(pp.buFont && !/^(Arial|Symbol|\+mn-lt|\+mj-lt)$/i.test(pp.buFont as string) ? { font: pp.buFont as string } : {}) };
  return b;
}

function cleanRun<T extends Omit<Run, 'text'>>(r: T): T {
  const out = { ...r } as Record<string, unknown>;
  for (const k of Object.keys(out)) if (out[k] === undefined || out[k] === false) delete out[k];
  if (!out.font) out.font = '+minor';
  return out as T;
}

/* ============================================================= geometry */

function customPaths(cg: Element): CustomPath[] {
  const out: CustomPath[] = [];
  const val = (v: string | null) => (v === null || Number.isNaN(Number(v)) ? 0 : Number(v));
  for (const p of kids(kid(cg, 'pathLst'), 'path')) {
    const cmds: PathCmd[] = [];
    for (const c of kids(p)) {
      const pts = kids(c, 'pt').flatMap((pt) => [val(attr(pt, 'x')), val(attr(pt, 'y'))]);
      switch (c.localName) {
        case 'moveTo':
          cmds.push({ c: 'M', p: pts });
          break;
        case 'lnTo':
          cmds.push({ c: 'L', p: pts });
          break;
        case 'cubicBezTo':
          cmds.push({ c: 'C', p: pts });
          break;
        case 'quadBezTo':
          cmds.push({ c: 'Q', p: pts });
          break;
        case 'arcTo':
          cmds.push({ c: 'A', p: [val(attr(c, 'wR')), val(attr(c, 'hR')), val(attr(c, 'stAng')) / 60000, val(attr(c, 'swAng')) / 60000] });
          break;
        case 'close':
          cmds.push({ c: 'Z', p: [] });
          break;
      }
    }
    out.push({ w: num(p, 'w'), h: num(p, 'h'), cmds, fill: attr(p, 'fill') !== 'none', stroke: attr(p, 'stroke') !== '0' && attr(p, 'stroke') !== 'false' });
  }
  return out;
}

interface Frame {
  x: number;
  y: number;
  w: number;
  h: number;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
}

function frameOf(xfrm: Element | null): Frame | null {
  if (!xfrm) return null;
  const off = kid(xfrm, 'off');
  const ext = kid(xfrm, 'ext');
  if (!off || !ext) return null;
  const f: Frame = { x: px(attr(off, 'x')), y: px(attr(off, 'y')), w: px(attr(ext, 'cx')), h: px(attr(ext, 'cy')) };
  const rot = num(xfrm, 'rot') / 60000;
  if (rot) f.rot = ((rot % 360) + 360) % 360;
  if (attr(xfrm, 'flipH') === '1') f.flipH = true;
  if (attr(xfrm, 'flipV') === '1') f.flipV = true;
  return f;
}

/** Maps child coordinates of a group into the group's parent space. */
interface GroupTransform {
  (f: Frame): Frame;
}

const identity: GroupTransform = (f) => f;

/* =========================================================== placeholder */

const PH_MAP: Record<string, PlaceholderType> = { title: 'title', ctrTitle: 'ctrTitle', subTitle: 'subTitle', body: 'body', obj: 'obj', pic: 'pic', dt: 'dt', ftr: 'ftr', sldNum: 'sldNum', tbl: 'obj', chart: 'obj', dgm: 'obj', media: 'obj', clipArt: 'pic' };

function phOf(nv: Element | null): { type: string; idx: string } | null {
  const ph = path(nv, 'nvPr', 'ph');
  if (!ph) return null;
  // a placeholder without a type is a content ("obj") placeholder
  return { type: attr(ph, 'type') ?? 'obj', idx: attr(ph, 'idx') ?? '0' };
}

function findPh(list: PhInfo[], ph: { type: string; idx: string }): PhInfo | undefined {
  const norm = (t: string) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
  return (
    list.find((p) => ph.idx !== '0' && p.idx === ph.idx) ??
    list.find((p) => p.type === ph.type && (ph.type === 'title' || ph.type === 'ctrTitle' || p.idx === ph.idx)) ??
    list.find((p) => p.type === ph.type) ??
    list.find((p) => norm(p.type) === norm(ph.type))
  );
}

function roleStyle(type: string | undefined, m: MasterInfo): ListStyle {
  if (type === 'title' || type === 'ctrTitle') return m.title;
  if (!type) return m.other;
  return m.body;
}

/* ============================================================== shapes */

interface ShapeScope {
  ctx: Ctx;
  env: ColorEnv;
  master: MasterInfo | null;
  layout: LayoutInfo | null;
  /** Text inherited for non-placeholder text (presentation defaults). */
  base: ListStyle;
  gt: GroupTransform;
  spidMap: Map<string, string>;
  /** Parsing a layout/master: placeholders become PhInfo, not elements. */
  collectPh?: PhInfo[];
}

function styleRefs(style: Element | null, sc: ShapeScope): { fill?: Fill; line?: Line | null; fontColor?: string } {
  if (!style) return {};
  const out: { fill?: Fill; line?: Line | null; fontColor?: string } = {};
  const fillRef = kid(style, 'fillRef');
  if (fillRef) {
    const idx = num(fillRef, 'idx');
    const c = colorOf(fillRef, sc.env);
    if (idx === 0) out.fill = { type: 'none' };
    else if (c) {
      const tmpl = idx >= 1000 ? sc.ctx.theme.bgFills[idx - 1001] : sc.ctx.theme.fills[idx - 1];
      const f = tmpl ? fillOf(tmpl, sc.ctx, { ...sc.env, ph: c }) : undefined;
      out.fill = f && f.type !== 'image' ? f : { type: 'solid', color: c };
    }
  }
  const lnRef = kid(style, 'lnRef');
  if (lnRef) {
    const idx = num(lnRef, 'idx');
    const c = colorOf(lnRef, sc.env);
    if (idx === 0 || !c) out.line = null;
    else {
      const tmpl = sc.ctx.theme.lines[idx - 1];
      out.line = tmpl ? (lineOf(tmpl, sc.ctx, { ...sc.env, ph: c }) ?? { color: c, width: 1 }) : { color: c, width: 1 };
    }
  }
  const fontRefEl = kid(style, 'fontRef');
  if (fontRefEl) out.fontColor = colorOf(fontRefEl, sc.env);
  return out;
}

function shapeFrom(el: Element, sc: ShapeScope): El | null {
  const nv = kid(el, el.localName === 'cxnSp' ? 'nvCxnSpPr' : 'nvSpPr');
  const cNv = kid(nv, 'cNvPr');
  const spPr = kid(el, 'spPr');
  const ph = phOf(nv);
  const phType = ph ? PH_MAP[ph.type] ?? 'body' : undefined;
  if (ph && sc.collectPh) {
    sc.collectPh.push({ type: ph.type, idx: ph.idx, xfrm: kid(spPr, 'xfrm'), body: kid(kid(el, 'txBody'), 'bodyPr'), lst: listStyle(path(el, 'txBody', 'lstStyle'), sc.ctx, sc.env), spPr });
    return null;
  }
  const layoutPh = ph && sc.layout ? findPh(sc.layout.phs, ph) : undefined;
  const masterPh = ph && sc.master ? findPh(sc.master.phs, { type: layoutPh?.type ?? ph.type, idx: layoutPh?.idx ?? ph.idx }) : undefined;
  const frame = frameOf(kid(spPr, 'xfrm')) ?? frameOf(layoutPh?.xfrm ?? null) ?? frameOf(masterPh?.xfrm ?? null);
  if (!frame) return null;
  const f = sc.gt(frame);
  const refs = styleRefs(kid(el, 'style'), sc);
  const prst = kid(spPr, 'prstGeom');
  const cust = kid(spPr, 'custGeom');
  const geom = prst ? (attr(prst, 'prst') ?? 'rect') : cust ? 'custom' : 'rect';
  const adj: Record<string, number> = {};
  for (const gd of kids(kid(prst, 'avLst'), 'gd')) {
    const m = /val\s+(-?\d+)/.exec(attr(gd, 'fmla') ?? '');
    if (m) adj[attr(gd, 'name') ?? 'adj'] = Number(m[1]);
  }
  const inheritedFill = fillOf(layoutPh?.spPr, sc.ctx, sc.env) ?? fillOf(masterPh?.spPr, sc.ctx, sc.env);
  const fill = fillOf(spPr, sc.ctx, sc.env) ?? refs.fill ?? inheritedFill;
  const lnEl = kid(spPr, 'ln');
  const line = lnEl ? lineOf(lnEl, sc.ctx, sc.env, refs.line ?? undefined) : refs.line !== undefined ? refs.line : (lineOf(kid(layoutPh?.spPr ?? null, 'ln'), sc.ctx, sc.env) ?? undefined);
  const inherited = ph && sc.master ? mergeLevels(sc.base, roleStyle(ph.type, sc.master), masterPh?.lst, layoutPh?.lst) : sc.base;
  const inheritedBody = { ...bodyProps(masterPh?.body ?? null), ...bodyProps(layoutPh?.body ?? null) };
  const text = textOf(kid(el, 'txBody'), inherited, inheritedBody, sc.ctx, sc.env, refs.fontColor ? { color: refs.fontColor } : {});
  const id = newId();
  if (cNv) sc.spidMap.set(attr(cNv, 'id') ?? '', id);
  const out: ShapeEl = {
    id,
    type: 'shape',
    name: attr(cNv, 'name') ?? undefined,
    geom,
    x: f.x,
    y: f.y,
    w: f.w,
    h: f.h,
    ...(f.rot ? { rot: f.rot } : {}),
    ...(f.flipH ? { flipH: true } : {}),
    ...(f.flipV ? { flipV: true } : {}),
    fill: fill ?? { type: 'none' },
    line: line ?? null,
  };
  if (Object.keys(adj).length) out.adj = adj;
  if (cust) out.custom = customPaths(cust);
  const shadow = shadowOf(spPr, sc.env);
  if (shadow) out.shadow = shadow;
  if (text) out.text = text;
  if (phType) {
    out.ph = phType;
    if (ph && ph.idx !== '0') out.phIdx = Number(ph.idx);
  }
  if (attr(kid(nv, 'cNvSpPr'), 'txBox') === '1') out.textbox = true;
  if (attr(cNv, 'descr')) out.alt = attr(cNv, 'descr')!;
  if (attr(cNv, 'hidden') === '1') out.hidden = true;
  const link = kid(cNv, 'hlinkClick');
  if (link) {
    const l = linkOf(link, sc.ctx);
    if (l) out.link = l;
  }
  return out;
}

function pictureFrom(el: Element, sc: ShapeScope): El | null {
  const nv = kid(el, 'nvPicPr');
  const cNv = kid(nv, 'cNvPr');
  const spPr = kid(el, 'spPr');
  const ph = phOf(nv);
  const layoutPh = ph && sc.layout ? findPh(sc.layout.phs, ph) : undefined;
  const frame = frameOf(kid(spPr, 'xfrm')) ?? frameOf(layoutPh?.xfrm ?? null);
  if (!frame) return null;
  const f = sc.gt(frame);
  const blipFill = kid(el, 'blipFill');
  const blip = kid(blipFill, 'blip');
  let src = mediaUrl(sc.ctx, rid(blip));
  let svg: string | undefined;
  const svgBlip = blip ? Array.from(blip.getElementsByTagNameNS('http://schemas.microsoft.com/office/drawing/2016/SVG/main', 'svgBlip'))[0] : undefined;
  if (svgBlip) {
    const rel = sc.ctx.rels.get(rid(svgBlip) ?? '');
    const data = rel ? sc.ctx.pkg.zip[rel.target] : undefined;
    if (data) svg = strFromU8(data);
  }
  if (!src && svg) src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  if (!src) {
    const video = path(nv, 'nvPr', 'videoFile') ?? path(nv, 'nvPr', 'audioFile');
    if (video) sc.ctx.warnings.add('Videos and audio are shown as their preview picture.');
    return null;
  }
  const rect = kid(blipFill, 'srcRect');
  const crop: [number, number, number, number] = [num(rect, 'l'), num(rect, 't'), num(rect, 'r'), num(rect, 'b')].map((v) => Math.max(0, Math.min(0.95, v / 100000))) as [number, number, number, number];
  const id = newId();
  if (cNv) sc.spidMap.set(attr(cNv, 'id') ?? '', id);
  const prst = attr(kid(spPr, 'prstGeom'), 'prst');
  const out: ImageEl = { id, type: 'image', name: attr(cNv, 'name') ?? 'Picture', src, x: f.x, y: f.y, w: f.w, h: f.h, ...(f.rot ? { rot: f.rot } : {}), ...(f.flipH ? { flipH: true } : {}), ...(f.flipV ? { flipV: true } : {}) };
  if (svg) out.svg = svg;
  if (crop.some((v) => v > 0.0005)) out.crop = crop;
  if (prst && prst !== 'rect') out.geom = prst;
  const line = lineOf(kid(spPr, 'ln'), sc.ctx, sc.env);
  if (line) out.line = line;
  const shadow = shadowOf(spPr, sc.env);
  if (shadow) out.shadow = shadow;
  if (attr(cNv, 'descr')) out.alt = attr(cNv, 'descr')!;
  if (ph) out.ph = PH_MAP[ph.type] ?? 'pic';
  const link = kid(cNv, 'hlinkClick');
  if (link) {
    const l = linkOf(link, sc.ctx);
    if (l) out.link = l;
  }
  return out;
}

function tableFrom(tbl: Element, f: Frame, name: string | undefined, sc: ShapeScope): TableEl {
  const tblPr = kid(tbl, 'tblPr');
  const styleId = kid(tblPr, 'tableStyleId')?.textContent?.trim();
  const style = { ...styleFromId(styleId), firstRow: attr(tblPr, 'firstRow') === '1', lastRow: attr(tblPr, 'lastRow') === '1', firstCol: attr(tblPr, 'firstCol') === '1', lastCol: attr(tblPr, 'lastCol') === '1', bandRows: attr(tblPr, 'bandRow') === '1', bandCols: attr(tblPr, 'bandCol') === '1' };
  const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => px(attr(g, 'w')));
  const base = sc.base;
  const rows = kids(tbl, 'tr').map((tr) => ({
    h: px(attr(tr, 'h')),
    cells: kids(tr, 'tc').map((tc): TableCell => {
      const tcPr = kid(tc, 'tcPr');
      const text = textOf(kid(tc, 'txBody'), base, {}, sc.ctx, sc.env) ?? { paras: [{ runs: [] }] };
      const anchor = attr(tcPr, 'anchor');
      text.anchor = anchor === 'ctr' ? 'm' : anchor === 'b' ? 'b' : 't';
      const mar = ['marL', 'marT', 'marR', 'marB'].map((k) => attr(tcPr, k));
      text.inset = [mar[0] !== null ? px(mar[0]) : 9.6, mar[1] !== null ? px(mar[1]) : 4.8, mar[2] !== null ? px(mar[2]) : 9.6, mar[3] !== null ? px(mar[3]) : 4.8];
      text.autofit = 'none';
      const cell: TableCell = { text };
      const fill = fillOf(tcPr ? (kids(tcPr).find((k) => ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill'].includes(k.localName)) ?? null) : null, sc.ctx, sc.env);
      if (fill) cell.fill = fill;
      const bl = (n: string) => {
        const e = kid(tcPr, n);
        return e ? (lineOf(e, sc.ctx, sc.env) ?? null) : undefined;
      };
      const borders = [bl('lnT'), bl('lnR'), bl('lnB'), bl('lnL')];
      if (borders.some((b) => b !== undefined)) cell.borders = borders.map((b) => b ?? null) as TableCell['borders'];
      if (num(tc, 'rowSpan', 1) > 1) cell.rowSpan = num(tc, 'rowSpan', 1);
      if (num(tc, 'gridSpan', 1) > 1) cell.colSpan = num(tc, 'gridSpan', 1);
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') cell.merged = true;
      return cell;
    }),
  }));
  return { id: newId(), type: 'table', name: name ?? 'Table', x: f.x, y: f.y, w: cols.reduce((a, b) => a + b, 0) || f.w, h: rows.reduce((a, r) => a + r.h, 0) || f.h, cols, rows, style };
}

function frameFrom(el: Element, sc: ShapeScope): El[] {
  const nv = kid(el, 'nvGraphicFramePr');
  const cNv = kid(nv, 'cNvPr');
  const frame = frameOf(kid(el, 'xfrm'));
  if (!frame) return [];
  const f = sc.gt(frame);
  const data = path(el, 'graphic', 'graphicData');
  const uri = attr(data, 'uri') ?? '';
  const name = attr(cNv, 'name') ?? undefined;
  const reg = (e: El) => {
    if (cNv) sc.spidMap.set(attr(cNv, 'id') ?? '', e.id);
    return e;
  };
  if (uri.endsWith('/table')) {
    const tbl = kid(data, 'tbl');
    return tbl ? [reg(tableFrom(tbl, f, name, sc))] : [];
  }
  if (uri.endsWith('/chart')) {
    const c = kid(data, 'chart');
    const rel = sc.ctx.rels.get(rid(c, 'id') ?? '');
    const xml = rel ? sc.ctx.pkg.text(rel.target) : null;
    const chart = xml ? parseChartXml(xml, (spPr) => fillOf(spPr ? (kids(spPr).find((k) => k.localName === 'solidFill') ?? null) : null, sc.ctx, sc.env)?.type === 'solid' ? (fillOf(kids(spPr!).find((k) => k.localName === 'solidFill') ?? null, sc.ctx, sc.env) as { color: string }).color : undefined) : null;
    if (chart) return [reg({ id: newId(), type: 'chart', name: name ?? 'Chart', x: f.x, y: f.y, w: f.w, h: f.h, chart })];
    return [];
  }
  if (uri.endsWith('/diagram')) {
    // SmartArt: use the pre-rendered drawing PowerPoint stores next to the diagram
    const relIds = kid(data, 'relIds');
    const dmRel = sc.ctx.rels.get(rid(relIds, 'dm') ?? '');
    let drawingPath: string | undefined;
    for (const r of sc.ctx.rels.values()) if (r.type === 'diagramDrawing') drawingPath = r.target;
    if (dmRel && !drawingPath) drawingPath = undefined;
    const doc = drawingPath ? sc.ctx.pkg.xml(drawingPath) : null;
    const tree = doc ? Array.from(doc.getElementsByTagNameNS('*', 'spTree'))[0] : null;
    if (!tree) {
      sc.ctx.warnings.add('Some SmartArt graphics could not be shown.');
      return [];
    }
    const savedRels = sc.ctx.rels;
    sc.ctx.rels = sc.ctx.pkg.rels(drawingPath!);
    const children = treeEls(tree, { ...sc, gt: (fr) => sc.gt({ ...fr, x: fr.x + frame.x, y: fr.y + frame.y }) });
    sc.ctx.rels = savedRels;
    if (!children.length) return [];
    const b = children.reduce((acc, c) => ({ x1: Math.min(acc.x1, c.x), y1: Math.min(acc.y1, c.y), x2: Math.max(acc.x2, c.x + c.w), y2: Math.max(acc.y2, c.y + c.h) }), { x1: Infinity, y1: Infinity, x2: -Infinity, y2: -Infinity });
    return [reg({ id: newId(), type: 'group', name: name ?? 'SmartArt', x: b.x1, y: b.y1, w: b.x2 - b.x1, h: b.y2 - b.y1, children } as GroupEl)];
  }
  // OLE objects and other frames: use the fallback picture when there is one
  const pic = Array.from(el.getElementsByTagNameNS('*', 'pic'))[0];
  if (pic) {
    const p = pictureFrom(pic, { ...sc, gt: () => f });
    return p ? [p] : [];
  }
  sc.ctx.warnings.add('Some embedded objects could not be shown.');
  return [];
}

function groupFrom(el: Element, sc: ShapeScope): El | null {
  const xfrm = path(el, 'grpSpPr', 'xfrm');
  const f = frameOf(xfrm);
  const chOff = kid(xfrm, 'chOff');
  const chExt = kid(xfrm, 'chExt');
  const cx = px(attr(chOff, 'x'));
  const cy = px(attr(chOff, 'y'));
  const cw = px(attr(chExt, 'cx')) || f?.w || 1;
  const ch = px(attr(chExt, 'cy')) || f?.h || 1;
  const outer = f ? sc.gt({ x: f.x, y: f.y, w: f.w, h: f.h }) : null;
  const sx = f && cw ? f.w / cw : 1;
  const sy = f && ch ? f.h / ch : 1;
  const inner: GroupTransform = (fr) => {
    const mapped = { ...fr, x: (f?.x ?? 0) + (fr.x - cx) * sx, y: (f?.y ?? 0) + (fr.y - cy) * sy, w: fr.w * sx, h: fr.h * sy };
    return sc.gt(mapped);
  };
  const children = treeEls(el, { ...sc, gt: inner });
  if (!children.length) return null;
  const cNv = path(el, 'nvGrpSpPr', 'cNvPr');
  const id = newId();
  if (cNv) sc.spidMap.set(attr(cNv, 'id') ?? '', id);
  const g: GroupEl = { id, type: 'group', name: attr(cNv, 'name') ?? 'Group', x: outer?.x ?? 0, y: outer?.y ?? 0, w: outer?.w ?? 0, h: outer?.h ?? 0, children };
  if (f?.rot) g.rot = f.rot;
  // group fill inheritance (grpFill)
  const gf = fillOf(path(el, 'grpSpPr'), sc.ctx, sc.env);
  if (gf) for (const c of children) if (c.type === 'shape' && !c.fill) c.fill = gf;
  return g;
}

function treeEls(tree: Element, sc: ShapeScope): El[] {
  const out: El[] = [];
  for (const raw of kids(tree)) {
    for (const el of unwrap(raw)) {
      let e: El | null = null;
      switch (el.localName) {
        case 'sp':
        case 'cxnSp':
          e = shapeFrom(el, sc);
          break;
        case 'pic':
          e = pictureFrom(el, sc);
          break;
        case 'grpSp':
          e = groupFrom(el, sc);
          break;
        case 'graphicFrame':
          out.push(...frameFrom(el, sc));
          break;
      }
      if (e) out.push(e);
    }
  }
  return out;
}

/* ================================================================== bg */

function backgroundOf(cSld: Element | null, ctx: Ctx, env: ColorEnv): Fill | undefined {
  const bg = kid(cSld, 'bg');
  if (!bg) return undefined;
  const bgPr = kid(bg, 'bgPr');
  if (bgPr) return fillOf(bgPr, ctx, env);
  const ref = kid(bg, 'bgRef');
  if (ref) {
    const idx = num(ref, 'idx');
    const c = colorOf(ref, env);
    const tmpl = idx >= 1001 ? ctx.theme.bgFills[idx - 1001] : ctx.theme.fills[idx - 1];
    const f = tmpl ? fillOf(tmpl, ctx, { ...env, ph: c }) : undefined;
    return f ?? (c ? { type: 'solid', color: c } : undefined);
  }
  return undefined;
}

/* ============================================================ animation */

const TRANS: Record<string, TransitionType> = { fade: 'fade', push: 'push', wipe: 'wipe', split: 'split', cover: 'cover', pull: 'reveal', reveal: 'reveal', zoom: 'zoom', circle: 'circle', dissolve: 'dissolve', flip: 'flip', morph: 'morph', cut: 'none', randomBar: 'wipe', blinds: 'wipe', checker: 'dissolve', comb: 'wipe', diamond: 'circle', plus: 'circle', wedge: 'circle', wheel: 'circle', newsflash: 'zoom', strips: 'wipe', vortex: 'zoom', ripple: 'circle', gallery: 'push', conveyor: 'push', pan: 'push', ferris: 'push', flash: 'fade', shred: 'dissolve', switch: 'flip', doors: 'split', window: 'split', box: 'zoom', warp: 'zoom', flythrough: 'zoom', glitter: 'dissolve', honeycomb: 'dissolve', prism: 'flip', curtains: 'split', wind: 'push', prestige: 'fade', fracture: 'dissolve', crush: 'zoom', peelOff: 'reveal', pageCurlDouble: 'reveal', pageCurlSingle: 'reveal', airplane: 'push', origami: 'flip' };

function transitionFrom(root: Element): Transition | undefined {
  let t: Element | null = null;
  for (const c of kids(root)) {
    if (c.localName === 'transition') t = c;
    if (c.localName === 'AlternateContent') for (const u of unwrap(c)) if (u.localName === 'transition') t = u;
  }
  if (!t) return undefined;
  const effect = kids(t).find((k) => TRANS[k.localName]);
  const type = effect ? TRANS[effect.localName] : 'none';
  const spd = attr(t, 'spd');
  const durAttr = Array.from(t.attributes).find((a) => a.localName === 'dur')?.value;
  const out: Transition = { type, dur: durAttr ? Number(durAttr) : spd === 'fast' ? 500 : spd === 'slow' ? 1000 : 750 };
  const dir = attr(effect, 'dir');
  if (dir && ['l', 'r', 'u', 'd'].includes(dir)) out.dir = dir as Dir;
  if (effect?.localName === 'split') out.dir = attr(effect, 'orient') === 'horz' ? 'u' : 'l';
  if (attr(t, 'advTm')) out.after = Number(attr(t, 'advTm'));
  if (attr(t, 'advClick') === '0') out.onClick = false;
  if (type === 'none' && !out.after) return undefined;
  return out;
}

const ENTR: Record<number, AnimEffect> = { 1: 'appear', 2: 'fly', 10: 'fade', 16: 'split', 21: 'wheel', 22: 'wipe', 23: 'zoom', 26: 'bounce', 30: 'float', 37: 'float', 42: 'float', 47: 'float', 53: 'zoom', 55: 'zoom' };
const EMPH: Record<number, AnimEffect> = { 6: 'grow', 8: 'spin', 9: 'transparency', 26: 'pulse', 32: 'teeter', 35: 'pulse' };
const EXIT: Record<number, AnimEffect> = { 1: 'disappear', 2: 'fly', 10: 'fade', 22: 'wipe', 23: 'zoom', 53: 'zoom' };
const SUB_DIR: Record<number, Dir> = { 1: 'u', 2: 'r', 4: 'd', 8: 'l' };

function animsFrom(root: Element, spidMap: Map<string, string>): Anim[] {
  const timing = kid(root, 'timing');
  if (!timing) return [];
  const seq = Array.from(timing.getElementsByTagNameNS(NS_P, 'cTn')).find((c) => attr(c, 'nodeType') === 'mainSeq');
  if (!seq) return [];
  const out: Anim[] = [];
  const effects = Array.from(seq.getElementsByTagNameNS(NS_P, 'cTn')).filter((c) => attr(c, 'presetClass'));
  for (const e of effects) {
    const cls = attr(e, 'presetClass');
    if (cls !== 'entr' && cls !== 'emph' && cls !== 'exit') continue;
    const id = num(e, 'presetID');
    const effect = (cls === 'entr' ? ENTR : cls === 'emph' ? EMPH : EXIT)[id] ?? (cls === 'emph' ? 'pulse' : 'fade');
    const tgt = Array.from(e.getElementsByTagNameNS(NS_P, 'spTgt'))[0];
    const el = spidMap.get(attr(tgt, 'spid') ?? '');
    if (!el) continue;
    const behaviours = Array.from(e.getElementsByTagNameNS(NS_P, 'cTn')).filter((c) => c !== e);
    const dur = Math.max(1, ...behaviours.map((c) => Number(attr(c, 'dur')) || 0).filter((v) => Number.isFinite(v)));
    const nodeType = attr(e, 'nodeType');
    const delay = Number(attr(path(e, 'stCondLst', 'cond'), 'delay') ?? 0) || 0;
    const para = tgt ? Array.from(tgt.getElementsByTagNameNS(NS_P, 'pRg'))[0] : undefined;
    const start: Anim['start'] = nodeType === 'clickEffect' ? 'click' : nodeType === 'afterEffect' ? 'after' : 'with';
    const prev = out[out.length - 1];
    if (para && prev && prev.el === el && prev.effect === effect && prev.byPara) continue;
    out.push({ id: newId('a'), el, cls: cls as AnimClass, effect, dir: SUB_DIR[num(e, 'presetSubtype')], start, dur: effect === 'appear' || effect === 'disappear' ? 1 : dur, delay, ...(para ? { byPara: true } : {}) });
  }
  return out;
}

/* ================================================================ main */

const LAYOUT_TYPES = new Set(['title', 'obj', 'secHead', 'twoObj', 'twoTxTwoObj', 'titleOnly', 'blank', 'objTx', 'picTx']);

function firstLevel(ls: ListStyle): LevelProps {
  return ls[0];
}

export async function importPptx(bytes: Uint8Array): Promise<LoadedPresentation> {
  const zip = unzipSync(bytes);
  const pkg = new Pkg(zip);
  const presDoc = pkg.xml('ppt/presentation.xml');
  if (!presDoc) throw new Error('This file is not a PowerPoint presentation.');
  const presRoot = presDoc.documentElement;
  const presRels = pkg.rels('ppt/presentation.xml');
  const sz = kid(presRoot, 'sldSz');
  const size = { w: Math.round(px(attr(sz, 'cx')) || 1280), h: Math.round(px(attr(sz, 'cy')) || 720) };
  const slidePaths = kids(kid(presRoot, 'sldIdLst'), 'sldId')
    .map((s) => presRels.get(rid(s, 'id') ?? '')?.target)
    .filter((p): p is string => !!p && pkg.has(p));

  // theme from the first master
  const masterPaths = kids(kid(presRoot, 'sldMasterIdLst'), 'sldMasterId')
    .map((m) => presRels.get(rid(m, 'id') ?? '')?.target)
    .filter((p): p is string => !!p && pkg.has(p));
  const firstMaster = masterPaths[0];
  const themePath = firstMaster ? [...pkg.rels(firstMaster).values()].find((r) => r.type === 'theme')?.target : undefined;
  const theme = readTheme(themePath ? pkg.xml(themePath) : null);
  const warnings = new Set<string>();
  const media = new Map<string, string>();
  const ctxFor = (partPath: string, env: ColorEnv): Ctx => ({ pkg, theme, env, partPath, rels: pkg.rels(partPath), slides: slidePaths, media, warnings });

  const defaultCtx = ctxFor('ppt/presentation.xml', { scheme: theme.scheme, map: {} });
  const defaults = listStyle(kid(presRoot, 'defaultTextStyle'), defaultCtx, defaultCtx.env);

  // masters
  const masters = new Map<string, MasterInfo>();
  const loadMaster = (p: string): MasterInfo => {
    const hit = masters.get(p);
    if (hit) return hit;
    const doc = pkg.xml(p)!;
    const root = doc.documentElement;
    const cm = kid(root, 'clrMap');
    const map: Record<string, string> = {};
    if (cm) for (const a of Array.from(cm.attributes)) map[a.localName] = a.value;
    const env: ColorEnv = { scheme: theme.scheme, map };
    const ctx = ctxFor(p, env);
    const tx = kid(root, 'txStyles');
    const m: MasterInfo = {
      path: p,
      env,
      title: mergeLevels(defaults, listStyle(kid(tx, 'titleStyle'), ctx, env)),
      body: mergeLevels(defaults, listStyle(kid(tx, 'bodyStyle'), ctx, env)),
      other: mergeLevels(defaults, listStyle(kid(tx, 'otherStyle'), ctx, env)),
      phs: [],
      bg: undefined,
      decor: [],
    };
    const cSld = kid(root, 'cSld');
    m.bg = backgroundOf(cSld, ctx, env);
    const sc: ShapeScope = { ctx, env, master: null, layout: null, base: defaults, gt: identity, spidMap: new Map(), collectPh: m.phs };
    m.decor = treeEls(kid(cSld, 'spTree')!, sc);
    masters.set(p, m);
    return m;
  };

  // layouts (only those used)
  const layouts = new Map<string, LayoutInfo>();
  const loadLayout = (p: string): LayoutInfo | null => {
    const hit = layouts.get(p);
    if (hit) return hit;
    const doc = pkg.xml(p);
    if (!doc) return null;
    const root = doc.documentElement;
    const mPath = [...pkg.rels(p).values()].find((r) => r.type === 'slideMaster')?.target ?? firstMaster;
    const master = mPath ? loadMaster(mPath) : null;
    if (!master) return null;
    const env = master.env;
    const ctx = ctxFor(p, env);
    const cSld = kid(root, 'cSld');
    const phs: PhInfo[] = [];
    const sc: ShapeScope = { ctx, env, master, layout: null, base: defaults, gt: identity, spidMap: new Map(), collectPh: phs };
    const decor = treeEls(kid(cSld, 'spTree')!, sc);
    const showMasterSp = attr(root, 'showMasterSp') !== '0';
    const type = attr(root, 'type') ?? 'cust';
    const info: LayoutInfo = {
      path: p,
      master,
      phs,
      showMasterSp,
      layout: {
        id: `layout-${layouts.size + 1}`,
        name: attr(cSld, 'name') ?? type,
        type: (LAYOUT_TYPES.has(type) ? type : 'cust') as LayoutType,
        background: backgroundOf(cSld, ctx, env),
        decor: [...(showMasterSp ? master.decor : []), ...decor],
        placeholders: [],
      },
    };
    // placeholders of the layout as elements for "New slide" with this layout
    const phScope: ShapeScope = { ctx, env, master, layout: { ...info, phs: [] }, base: defaults, gt: identity, spidMap: new Map() };
    const phEls: El[] = [];
    for (const raw of kids(kid(cSld, 'spTree'))) {
      for (const el of unwrap(raw)) {
        if (el.localName !== 'sp' || !phOf(kid(el, 'nvSpPr'))) continue;
        const ph = phOf(kid(el, 'nvSpPr'))!;
        if (['dt', 'ftr', 'sldNum'].includes(ph.type)) continue;
        const clone = el.cloneNode(true) as Element;
        // empty the prompt text but keep formatting
        for (const r of Array.from(clone.getElementsByTagNameNS('*', 'r'))) r.remove();
        const shape = shapeFrom(clone, { ...phScope, layout: info });
        if (shape) phEls.push(shape);
      }
    }
    info.layout.placeholders = phEls;
    layouts.set(p, info);
    return info;
  };

  const slides: Slide[] = [];
  let footer: Presentation['footer'] | undefined;
  for (let i = 0; i < slidePaths.length; i++) {
    const sp = slidePaths[i];
    const doc = pkg.xml(sp);
    if (!doc) continue;
    const root = doc.documentElement;
    const rels = pkg.rels(sp);
    const lPath = [...rels.values()].find((r) => r.type === 'slideLayout')?.target;
    const layout = lPath ? loadLayout(lPath) : null;
    const master = layout?.master ?? (firstMaster ? loadMaster(firstMaster) : null);
    const env: ColorEnv = { scheme: theme.scheme, map: master?.env.map ?? {} };
    const ovr = path(root, 'clrMapOvr', 'overrideClrMapping');
    if (ovr) for (const a of Array.from(ovr.attributes)) env.map[a.localName] = a.value;
    const ctx = ctxFor(sp, env);
    const spidMap = new Map<string, string>();
    const cSld = kid(root, 'cSld');
    const sc: ShapeScope = { ctx, env, master, layout, base: defaults, gt: identity, spidMap };
    let elements = treeEls(kid(cSld, 'spTree')!, sc);
    // footers become presentation-wide settings
    const fEls = elements.filter((e) => e.ph === 'dt' || e.ph === 'ftr' || e.ph === 'sldNum');
    if (fEls.length) {
      footer ??= {};
      for (const e of fEls) {
        if (e.ph === 'sldNum') footer.slideNumber = true;
        if (e.ph === 'dt') footer.date = true;
        if (e.ph === 'ftr' && e.type === 'shape' && e.text) footer.text = e.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join(' ').trim() || footer.text;
      }
      elements = elements.filter((e) => !fEls.includes(e));
    }
    const slide: Slide = { id: newId('s'), layout: layout?.layout.id ?? 'layout-1', elements };
    const bg = backgroundOf(cSld, ctx, env);
    if (bg) slide.background = bg;
    if (attr(root, 'show') === '0') slide.hidden = true;
    if (attr(root, 'showMasterSp') === '0') slide.hideDecor = true;
    const t = transitionFrom(root);
    if (t) slide.transition = t;
    const anims = animsFrom(root, spidMap);
    if (anims.length) slide.anims = anims;
    // speaker notes
    const nPath = [...rels.values()].find((r) => r.type === 'notesSlide')?.target;
    const nDoc = nPath ? pkg.xml(nPath) : null;
    if (nDoc) {
      const body = Array.from(nDoc.getElementsByTagNameNS(NS_P, 'sp')).find((s) => {
        const ph = path(s, 'nvSpPr', 'nvPr', 'ph');
        return ph && (attr(ph, 'type') === 'body' || (!attr(ph, 'type') && attr(ph, 'idx') === '1'));
      });
      const txt = body ? kids(kid(body, 'txBody'), 'p').map((p) => Array.from(p.getElementsByTagNameNS('*', 't')).map((t) => t.textContent ?? '').join('')).join('\n') : '';
      if (txt.trim()) slide.notes = txt.replace(/\s+$/, '');
    }
    slides.push(slide);
  }

  const m0 = firstMaster ? loadMaster(firstMaster) : null;
  const colors: Record<ThemeColorName, string> = {
    bg1: theme.scheme[m0?.env.map.bg1 ?? 'lt1'] ?? '#FFFFFF',
    tx1: theme.scheme[m0?.env.map.tx1 ?? 'dk1'] ?? '#000000',
    bg2: theme.scheme[m0?.env.map.bg2 ?? 'lt2'] ?? '#E7E6E6',
    tx2: theme.scheme[m0?.env.map.tx2 ?? 'dk2'] ?? '#44546A',
    accent1: theme.scheme.accent1,
    accent2: theme.scheme.accent2,
    accent3: theme.scheme.accent3,
    accent4: theme.scheme.accent4,
    accent5: theme.scheme.accent5,
    accent6: theme.scheme.accent6,
    hlink: theme.scheme.hlink,
    folHlink: theme.scheme.folHlink,
  };
  const th: Theme = { id: 'imported', name: theme.name, colors, fonts: { major: theme.major, minor: theme.minor } };
  const tl = m0 ? firstLevel(m0.title) : { p: {}, r: {} };
  const bl = m0 ? firstLevel(m0.body) : { p: {}, r: {} };
  const ol = m0 ? firstLevel(m0.other) : { p: {}, r: {} };
  const pres: Presentation = {
    format: 'affice-slides',
    version: 1,
    size,
    theme: th,
    master: {
      background: m0?.bg ?? { type: 'solid', color: '@bg1' },
      decor: [],
      text: {
        title: { font: tl.r.font ?? '+major', size: tl.r.size ?? 44, color: tl.r.color ?? '@tx1' },
        body: { font: bl.r.font ?? '+minor', size: bl.r.size ?? 28, color: bl.r.color ?? '@tx1', lineSpacing: (bl.p.lineSpacing as number) ?? 0.9 },
        other: { font: ol.r.font ?? '+minor', size: ol.r.size ?? 18, color: ol.r.color ?? '@tx1' },
      },
    },
    layouts: [...layouts.values()].map((l) => l.layout),
    slides: slides.length ? slides : [{ id: newId('s'), layout: 'layout-1', elements: [] }],
    props: {},
    footer,
  };
  if (!pres.layouts.length) pres.layouts = [{ id: 'layout-1', name: 'Blank', type: 'blank', decor: m0?.decor ?? [], placeholders: [] }];
  // layouts that no slide used still need to exist for "New slide": load the rest of the master's layouts
  if (firstMaster) {
    for (const r of pkg.rels(firstMaster).values()) if (r.type === 'slideLayout' && !layouts.has(r.target)) loadLayout(r.target);
    pres.layouts = [...layouts.values()].map((l) => l.layout);
  }
  // core properties
  const core = pkg.xml('docProps/core.xml');
  if (core) {
    const t = core.getElementsByTagNameNS('http://purl.org/dc/elements/1.1/', 'title')[0]?.textContent;
    const a = core.getElementsByTagNameNS('http://purl.org/dc/elements/1.1/', 'creator')[0]?.textContent;
    if (t) pres.props.title = t;
    if (a) pres.props.author = a;
  }
  return { pres, warning: warnings.size ? [...warnings].join(' ') : undefined };
}
