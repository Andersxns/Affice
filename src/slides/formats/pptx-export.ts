/**
 * PowerPoint (.pptx) writer. Produces a real theme, slide master with text styles, one layout per slide
 * layout (with decorative shapes and placeholders), slides, notes, native charts (with an embedded
 * workbook), transitions and animations using PowerPoint's own presets.
 */
import { zipSync, strToU8 } from 'fflate';
import { escapeXml as x } from '@/lib/utils';
import { DEFAULT_ADJ, isConnector } from '../geometry';
import {
  parseColorRef,
  type Anim,
  type ChartEl,
  type El,
  type Fill,
  type GroupEl,
  type ImageEl,
  type Layout,
  type Line,
  type Para,
  type Presentation,
  type Run,
  type Shadow,
  type ShapeEl,
  type Slide,
  type TableEl,
  type TextBody,
  type Transition,
} from '../model';
import { effectivePara, effectiveRun, roleOf, type TextContext } from '../render/text';
import { cellLook, tableStyleId } from '../tables';
import { buildOf } from '../show/engine';
import { decodeImage as decode, rasterizeImage as rasterize } from './media';
import { chartSpaceXml, chartWorkbook } from './pptx-charts';

const EMU = 9525;
const emu = (px: number) => Math.round(px * EMU);
const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/* ================================================================ colours */

export function colorXml(c: string | undefined): string {
  if (!c) return '<a:schemeClr val="tx1"/>';
  const ref = parseColorRef(c);
  if (ref) {
    const name = ref.name === 'dk1' ? 'tx1' : ref.name === 'lt1' ? 'bg1' : ref.name === 'dk2' ? 'tx2' : ref.name === 'lt2' ? 'bg2' : ref.name;
    let mods = '';
    if (ref.lum > 0) mods += `<a:lumMod val="${(100 - ref.lum) * 1000}"/><a:lumOff val="${ref.lum * 1000}"/>`;
    else if (ref.lum < 0) mods += `<a:lumMod val="${(100 + ref.lum) * 1000}"/>`;
    if (ref.alpha < 1) mods += `<a:alpha val="${Math.round(ref.alpha * 100000)}"/>`;
    return `<a:schemeClr val="${name}">${mods}</a:schemeClr>`;
  }
  let hex = c.trim();
  let alpha = 1;
  const rgba = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+))?\s*\)$/i.exec(hex);
  if (rgba) {
    hex = `#${[rgba[1], rgba[2], rgba[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('')}`;
    if (rgba[4] !== undefined) alpha = Number(rgba[4]);
  }
  let h = hex.replace('#', '');
  if (h.length === 3) h = [...h].map((ch) => ch + ch).join('');
  if (h.length === 8) {
    alpha = parseInt(h.slice(6, 8), 16) / 255;
    h = h.slice(0, 6);
  }
  if (!/^[0-9a-f]{6}$/i.test(h)) h = '000000';
  return `<a:srgbClr val="${h.toUpperCase()}">${alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ''}</a:srgbClr>`;
}

/* ================================================================ package */

interface Part {
  path: string;
  data: string | Uint8Array;
}

class Rels {
  items: { id: string; type: string; target: string; external?: boolean }[] = [];
  add(type: string, target: string, external = false): string {
    const found = this.items.find((i) => i.type === type && i.target === target && !!i.external === external);
    if (found) return found.id;
    const id = `rId${this.items.length + 1}`;
    this.items.push({ id, type, target, external });
    return id;
  }
  xml(): string {
    return `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${this.items.map((i) => `<Relationship Id="${i.id}" Type="${REL}/${i.type}" Target="${x(i.target)}"${i.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
  }
}

class Media {
  files: { path: string; data: Uint8Array; ext: string }[] = [];
  private bySrc = new Map<string, string>();
  async add(src: string): Promise<string | null> {
    const hit = this.bySrc.get(src);
    if (hit) return hit;
    const blob = await decode(src);
    if (!blob) return null;
    const path = `ppt/media/image${this.files.length + 1}.${blob.ext}`;
    this.files.push({ path, data: blob.data, ext: blob.ext });
    this.bySrc.set(src, path);
    return path;
  }
  async addBytes(data: Uint8Array, ext: string): Promise<string> {
    const path = `ppt/media/image${this.files.length + 1}.${ext}`;
    this.files.push({ path, data, ext });
    return path;
  }
}

/* ================================================================== fills */

function fillXml(f: Fill | undefined | null, blipRid?: string): string {
  if (!f || f.type === 'none') return '<a:noFill/>';
  if (f.type === 'solid') return `<a:solidFill>${colorXml(f.color)}</a:solidFill>`;
  if (f.type === 'gradient') {
    const stops = [...f.stops].sort((a, b) => a.pos - b.pos).map((s) => `<a:gs pos="${Math.round(s.pos * 100000)}">${colorXml(s.color)}</a:gs>`).join('');
    const shade = f.radial ? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' : `<a:lin ang="${Math.round((((f.angle % 360) + 360) % 360) * 60000)}" scaled="0"/>`;
    return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst>${shade}</a:gradFill>`;
  }
  if (!blipRid) return '<a:noFill/>';
  const mode = f.type === 'image' && f.mode === 'tile' ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>' : '<a:stretch><a:fillRect/></a:stretch>';
  return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${blipRid}"/><a:srcRect/>${mode}</a:blipFill>`;
}

const DASH: Record<string, string> = { dash: 'dash', dot: 'sysDot', sysDot: 'sysDot', sysDash: 'sysDash', dashDot: 'dashDot', longDash: 'lgDash' };
const ARROW: Record<string, string> = { triangle: 'triangle', arrow: 'arrow', stealth: 'stealth', oval: 'oval', diamond: 'diamond' };

function lineXml(l: Line | null | undefined): string {
  if (!l || l.width <= 0) return '<a:ln><a:noFill/></a:ln>';
  const cap = l.cap === 'round' ? ' cap="rnd"' : l.cap === 'square' ? ' cap="sq"' : '';
  const dash = l.dash && l.dash !== 'solid' ? `<a:prstDash val="${DASH[l.dash] ?? 'dash'}"/>` : '';
  const head = l.head && l.head !== 'none' ? `<a:headEnd type="${ARROW[l.head]}"/>` : '';
  const tail = l.tail && l.tail !== 'none' ? `<a:tailEnd type="${ARROW[l.tail]}"/>` : '';
  return `<a:ln w="${emu(l.width)}"${cap}><a:solidFill>${colorXml(l.color)}</a:solidFill>${dash}<a:round/>${head}${tail}</a:ln>`;
}

function shadowXml(s: Shadow | undefined): string {
  if (!s) return '';
  return `<a:effectLst><a:outerShdw blurRad="${emu(s.blur)}" dist="${emu(s.dist)}" dir="${Math.round(s.angle * 60000)}" algn="ctr" rotWithShape="0">${colorXml(s.color)}</a:outerShdw></a:effectLst>`;
}

function xfrm(e: El, tag = 'a:xfrm', child?: boolean): string {
  const rot = e.rot ? ` rot="${Math.round(e.rot * 60000)}"` : '';
  const flip = `${e.flipH ? ' flipH="1"' : ''}${e.flipV ? ' flipV="1"' : ''}`;
  const ch = child ? `<a:chOff x="${emu(e.x)}" y="${emu(e.y)}"/><a:chExt cx="${emu(Math.max(0, e.w))}" cy="${emu(Math.max(0, e.h))}"/>` : '';
  return `<${tag}${rot}${flip}><a:off x="${emu(e.x)}" y="${emu(e.y)}"/><a:ext cx="${emu(Math.max(0, e.w))}" cy="${emu(Math.max(0, e.h))}"/>${ch}</${tag}>`;
}

function geomXml(e: ShapeEl): string {
  if (e.geom === 'custom' && e.custom?.length) {
    const paths = e.custom
      .map((p) => {
        const cmds = p.cmds
          .map((c) => {
            const q = c.p.map((v) => Math.round(v));
            switch (c.c) {
              case 'M':
                return `<a:moveTo><a:pt x="${q[0]}" y="${q[1]}"/></a:moveTo>`;
              case 'L':
                return `<a:lnTo><a:pt x="${q[0]}" y="${q[1]}"/></a:lnTo>`;
              case 'C':
                return `<a:cubicBezTo><a:pt x="${q[0]}" y="${q[1]}"/><a:pt x="${q[2]}" y="${q[3]}"/><a:pt x="${q[4]}" y="${q[5]}"/></a:cubicBezTo>`;
              case 'Q':
                return `<a:quadBezTo><a:pt x="${q[0]}" y="${q[1]}"/><a:pt x="${q[2]}" y="${q[3]}"/></a:quadBezTo>`;
              case 'A':
                return `<a:arcTo wR="${q[0]}" hR="${q[1]}" stAng="${Math.round(c.p[2] * 60000)}" swAng="${Math.round(c.p[3] * 60000)}"/>`;
              default:
                return '<a:close/>';
            }
          })
          .join('');
        return `<a:path w="${Math.round(p.w)}" h="${Math.round(p.h)}"${p.fill === false ? ' fill="none"' : ''}${p.stroke === false ? ' stroke="0"' : ''}>${cmds}</a:path>`;
      })
      .join('');
    return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`;
  }
  const geom = e.geom === 'textBox' ? 'rect' : e.geom;
  const adj = e.adj ? Object.entries(e.adj).filter(([k, v]) => DEFAULT_ADJ[geom]?.[k] !== v) : [];
  return `<a:prstGeom prst="${x(geom)}"><a:avLst>${adj.map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join('')}</a:avLst></a:prstGeom>`;
}

/* =================================================================== text */

interface SlideCtx {
  rels: Rels;
  media: Media;
  pres: Presentation;
  slideIndex: number;
  /** Element id → numeric shape id. */
  ids: Map<string, number>;
  nextId: number;
  slidePath: (index: number) => string;
  charts: { n: number };
  parts: Part[];
  ct: Set<string>;
}

function fontXml(font: string | undefined): string {
  if (!font) return '';
  const f = font === '+major' ? '+mj-lt' : font === '+minor' ? '+mn-lt' : font;
  return `<a:latin typeface="${x(f)}"/><a:cs typeface="${x(f)}"/>`;
}

function rPrXml(r: Omit<Run, 'text'>, tag: 'a:rPr' | 'a:endParaRPr' | 'a:defRPr', ctx?: SlideCtx): string {
  const at: string[] = [];
  if (tag !== 'a:defRPr') at.push('lang="en-US"');
  if (r.size) at.push(`sz="${Math.round(r.size * 100)}"`);
  if (r.b !== undefined) at.push(`b="${r.b ? 1 : 0}"`);
  if (r.i !== undefined) at.push(`i="${r.i ? 1 : 0}"`);
  if (r.u) at.push('u="sng"');
  if (r.s) at.push('strike="sngStrike"');
  if (r.sup) at.push('baseline="30000"');
  if (r.sub) at.push('baseline="-25000"');
  if (r.caps) at.push('cap="all"');
  if (r.spacing) at.push(`spc="${Math.round(r.spacing * 100)}"`);
  if (tag !== 'a:defRPr') at.push('dirty="0"');
  let kids = '';
  if (r.color) kids += `<a:solidFill>${colorXml(r.color)}</a:solidFill>`;
  if (r.hl) kids += `<a:highlight>${colorXml(r.hl)}</a:highlight>`;
  kids += fontXml(r.font);
  if (r.link && ctx) kids += linkXml(r.link, ctx);
  return kids ? `<${tag} ${at.join(' ')}>${kids}</${tag}>` : `<${tag} ${at.join(' ')}/>`;
}

function linkXml(href: string, ctx: SlideCtx, tag = 'a:hlinkClick'): string {
  const jump: Record<string, string> = { '#next': 'nextslide', '#prev': 'previousslide', '#first': 'firstslide', '#last': 'lastslide' };
  if (jump[href]) return `<${tag} r:id="" action="ppaction://hlinkshowjump?jump=${jump[href]}"/>`;
  const m = /^#slide:(\d+)$/.exec(href);
  if (m) {
    const idx = Number(m[1]) - 1;
    if (idx >= 0 && idx < ctx.pres.slides.length) {
      const rid = ctx.rels.add('slide', `../slides/slide${idx + 1}.xml`);
      return `<${tag} r:id="${rid}" action="ppaction://hlinksldjump"/>`;
    }
    return '';
  }
  const rid = ctx.rels.add('hyperlink', href, true);
  return `<${tag} r:id="${rid}"/>`;
}

const BU_NUM: Record<string, string> = { arabicPeriod: 'arabicPeriod', arabicParenR: 'arabicParenR', alphaLcPeriod: 'alphaLcPeriod', alphaUcPeriod: 'alphaUcPeriod', alphaLcParenR: 'alphaLcParenR', romanLcPeriod: 'romanLcPeriod', romanUcPeriod: 'romanUcPeriod' };

function pPrXml(p: Para, tctx: TextContext | null, level = p.level ?? 0): string {
  const at: string[] = [];
  const e = tctx ? effectivePara(p, tctx) : null;
  // write the effective indents whenever they differ from what the master's list style would give
  const own = !!e && (p.bullet !== undefined || tctx?.role !== 'body');
  if (p.marL !== undefined) at.push(`marL="${emu(p.marL)}"`);
  else if (own) at.push(`marL="${emu(e.marL)}"`);
  if (p.indent !== undefined) at.push(`indent="${emu(p.indent)}"`);
  else if (own) at.push(`indent="${emu(e.indent)}"`);
  if (level) at.push(`lvl="${level}"`);
  if (p.align && p.align !== 'left') at.push(`algn="${p.align === 'center' ? 'ctr' : p.align === 'right' ? 'r' : 'just'}"`);
  else if (p.align === 'left') at.push('algn="l"');
  let kids = '';
  if (p.lineSpacing !== undefined) kids += `<a:lnSpc><a:spcPct val="${Math.round(p.lineSpacing * 100000)}"/></a:lnSpc>`;
  if (p.spaceBefore !== undefined) kids += `<a:spcBef><a:spcPts val="${Math.round(p.spaceBefore * 100)}"/></a:spcBef>`;
  if (p.spaceAfter !== undefined) kids += `<a:spcAft><a:spcPts val="${Math.round(p.spaceAfter * 100)}"/></a:spcAft>`;
  const b = p.bullet;
  if (b?.type === 'none') kids += '<a:buNone/>';
  else if (b?.type === 'char') kids += `${b.color ? `<a:buClr>${colorXml(b.color)}</a:buClr>` : ''}<a:buFont typeface="${x(b.font ?? 'Arial')}"/><a:buChar char="${x(b.char)}"/>`;
  else if (b?.type === 'num') kids += `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${BU_NUM[b.style] ?? 'arabicPeriod'}"${b.start && b.start !== 1 ? ` startAt="${b.start}"` : ''}/>`;
  if (!at.length && !kids) return '';
  return kids ? `<a:pPr ${at.join(' ')}>${kids}</a:pPr>` : `<a:pPr ${at.join(' ')}/>`;
}

function parasXml(body: TextBody, tctx: TextContext | null, ctx: SlideCtx, extraRun?: Omit<Run, 'text'>): string {
  const paras = body.paras.length ? body.paras : [{ runs: [] }];
  return paras
    .map((p) => {
      let runs = '';
      for (const r0 of p.runs) {
        const r = extraRun ? { ...extraRun, ...r0 } : r0;
        if (r.field) {
          const type = r.field === 'slidenum' ? 'slidenum' : 'datetime1';
          runs += `<a:fld id="{${guid(`${ctx.slideIndex}-${type}-${runs.length}`)}}" type="${type}">${rPrXml(r, 'a:rPr', ctx)}<a:t>${x(r.text)}</a:t></a:fld>`;
          continue;
        }
        const parts = r.text.split('\n');
        parts.forEach((t, i) => {
          if (i > 0) runs += `<a:br>${rPrXml({ ...r, link: undefined }, 'a:rPr')}</a:br>`;
          if (t) runs += `<a:r>${rPrXml(r, 'a:rPr', ctx)}<a:t>${x(t)}</a:t></a:r>`;
        });
      }
      const end = p.endRun ?? p.runs[p.runs.length - 1] ?? {};
      const endRun: Omit<Run, 'text'> = extraRun ? { ...extraRun, ...end } : { ...end };
      delete (endRun as Partial<Run>).text;
      delete endRun.link;
      delete endRun.field;
      return `<a:p>${pPrXml(p, tctx)}${runs}${rPrXml(endRun, 'a:endParaRPr')}</a:p>`;
    })
    .join('');
}

function bodyPrXml(body: TextBody | undefined, textbox: boolean): string {
  const at: string[] = ['rtlCol="0"'];
  const b = body ?? { paras: [] };
  at.push(`wrap="${b.wrap === false ? 'none' : 'square'}"`);
  if (b.inset) at.push(`lIns="${emu(b.inset[0])}" tIns="${emu(b.inset[1])}" rIns="${emu(b.inset[2])}" bIns="${emu(b.inset[3])}"`);
  const anchor = b.anchor ?? (textbox ? 't' : undefined);
  if (anchor) at.push(`anchor="${anchor === 'm' ? 'ctr' : anchor === 'b' ? 'b' : 't'}"`);
  if (b.vert && b.vert !== 'horz') at.push(`vert="${b.vert}"`);
  if (b.columns && b.columns > 1) at.push(`numCol="${b.columns}" spcCol="${emu(24)}"`);
  const fit = b.autofit === 'shrink' ? (b.fontScale && b.fontScale < 1 ? `<a:normAutofit fontScale="${Math.round(b.fontScale * 100000)}"/>` : '<a:normAutofit/>') : b.autofit === 'resize' ? '<a:spAutoFit/>' : '<a:noAutofit/>';
  return `<a:bodyPr ${at.join(' ')}>${fit}</a:bodyPr>`;
}

function lstStyleXml(defaults: TextBody['defaults'], levels = 1): string {
  if (!defaults || !Object.keys(defaults).length) return '<a:lstStyle/>';
  const def = rPrXml(defaults, 'a:defRPr');
  let out = '';
  for (let i = 1; i <= levels; i++) out += `<a:lvl${i}pPr>${def}</a:lvl${i}pPr>`;
  return `<a:lstStyle>${out}</a:lstStyle>`;
}

function guid(seed: string): string {
  let h1 = 0x811c9dc5;
  let h2 = 0x1234567;
  for (let i = 0; i < seed.length; i++) {
    h1 = Math.imul(h1 ^ seed.charCodeAt(i), 16777619);
    h2 = Math.imul(h2 ^ seed.charCodeAt(i), 2246822519);
  }
  const hex = (n: number) => (n >>> 0).toString(16).padStart(8, '0').toUpperCase();
  const a = hex(h1);
  const b = hex(h2);
  const c = hex(h1 ^ h2);
  return `${a}-${b.slice(0, 4)}-4${b.slice(5, 8)}-A${c.slice(1, 4)}-${c.slice(4)}${a.slice(0, 4)}`;
}

/* =============================================================== elements */

const PH_TYPE: Record<string, string> = { title: 'title', ctrTitle: 'ctrTitle', subTitle: 'subTitle', body: 'body', obj: 'obj', pic: 'pic', dt: 'dt', ftr: 'ftr', sldNum: 'sldNum' };

function phXml(e: El): string {
  if (!e.ph) return '<p:nvPr/>';
  // "obj" (content) is the default placeholder type and is written without a type, like PowerPoint does
  const type = e.ph === 'obj' ? '' : ` type="${PH_TYPE[e.ph]}"`;
  const idx = e.phIdx ? ` idx="${e.phIdx}"` : '';
  return `<p:nvPr><p:ph${type}${idx}/></p:nvPr>`;
}

function cNvPr(e: El, ctx: SlideCtx): string {
  const id = ctx.nextId++;
  ctx.ids.set(e.id, id);
  const name = x(e.name ?? `${e.type} ${id}`);
  const descr = e.alt ? ` descr="${x(e.alt)}"` : '';
  const hidden = e.hidden ? ' hidden="1"' : '';
  const link = e.link ? linkXml(e.link, ctx) : '';
  return `<p:cNvPr id="${id}" name="${name}"${descr}${hidden}>${link}</p:cNvPr>`;
}

async function shapeXml(e: ShapeEl, ctx: SlideCtx): Promise<string> {
  const connector = isConnector(e.geom) && !e.text;
  let blip: string | undefined;
  if (e.fill?.type === 'image') {
    const p = await ctx.media.add(e.fill.src);
    if (p) blip = ctx.rels.add('image', `../media/${p.split('/').pop()}`);
  }
  const spPr = `<p:spPr>${xfrm(e)}${geomXml(e)}${e.fill !== undefined ? fillXml(e.fill, blip) : e.textbox || e.ph ? '<a:noFill/>' : '<a:noFill/>'}${e.line !== undefined ? lineXml(e.line) : '<a:ln><a:noFill/></a:ln>'}${shadowXml(e.shadow)}</p:spPr>`;
  if (connector) return `<p:cxnSp><p:nvCxnSpPr>${cNvPr(e, ctx)}<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>${spPr}</p:cxnSp>`;
  const tctx: TextContext = { pres: ctx.pres, role: roleOf(e.ph), ph: e.ph, slideNumber: ctx.slideIndex + 1 };
  const body = e.text;
  const txBody = body ? `<p:txBody>${bodyPrXml(body, !!e.textbox)}${lstStyleXml(body.defaults, 9)}${parasXml(body, tctx, ctx)}</p:txBody>` : e.ph ? '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>' : '';
  return `<p:sp><p:nvSpPr>${cNvPr(e, ctx)}<p:cNvSpPr${e.textbox ? ' txBox="1"' : ''}/>${phXml(e)}</p:nvSpPr>${spPr}${txBody}</p:sp>`;
}

async function pictureXml(e: ImageEl, ctx: SlideCtx): Promise<string> {
  let rid = '';
  let svgRid = '';
  if (e.svg) {
    const svgPath = await ctx.media.addBytes(strToU8(e.svg), 'svg');
    svgRid = ctx.rels.add('image', `../media/${svgPath.split('/').pop()}`);
    const png = await rasterize(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(e.svg)}`, Math.max(64, e.w * 2), Math.max(64, e.h * 2));
    if (png) {
      const p = await ctx.media.addBytes(png, 'png');
      rid = ctx.rels.add('image', `../media/${p.split('/').pop()}`);
    }
  } else {
    const p = await ctx.media.add(e.src);
    if (p) rid = ctx.rels.add('image', `../media/${p.split('/').pop()}`);
  }
  if (!rid && !svgRid) return '';
  const ext = svgRid ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${svgRid}"/></a:ext></a:extLst>` : '';
  const [l, t, r, b] = e.crop ?? [0, 0, 0, 0];
  const src = e.crop ? `<a:srcRect l="${Math.round(l * 100000)}" t="${Math.round(t * 100000)}" r="${Math.round(r * 100000)}" b="${Math.round(b * 100000)}"/>` : '<a:srcRect/>';
  const geom = e.geom ?? 'rect';
  return `<p:pic><p:nvPicPr>${cNvPr(e, ctx)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>${phXml(e)}</p:nvPicPr><p:blipFill><a:blip r:embed="${rid || svgRid}">${ext}</a:blip>${src}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(e)}<a:prstGeom prst="${x(geom)}"><a:avLst/></a:prstGeom>${e.line ? lineXml(e.line) : ''}${shadowXml(e.shadow)}</p:spPr></p:pic>`;
}

function cellBorder(tag: string, l: Line | null | undefined): string {
  if (!l || l.width <= 0) return `<a:${tag} w="0"><a:noFill/></a:${tag}>`;
  return `<a:${tag} w="${emu(l.width)}"><a:solidFill>${colorXml(l.color)}</a:solidFill>${l.dash && l.dash !== 'solid' ? `<a:prstDash val="${DASH[l.dash] ?? 'dash'}"/>` : ''}</a:${tag}>`;
}

function tableXml(e: TableEl, ctx: SlideCtx): string {
  const st = e.style ?? {};
  const styleId = tableStyleId(st);
  const flags = `${st.firstRow ? ' firstRow="1"' : ''}${st.lastRow ? ' lastRow="1"' : ''}${st.firstCol ? ' firstCol="1"' : ''}${st.lastCol ? ' lastCol="1"' : ''}${st.bandRows ? ' bandRow="1"' : ''}${st.bandCols ? ' bandCol="1"' : ''}`;
  const tctx: TextContext = { pres: ctx.pres, role: 'other', slideNumber: ctx.slideIndex + 1 };
  const rows = e.rows
    .map((row, r) => {
      const cells = row.cells
        .map((cell, c) => {
          const look = cellLook(e, r, c);
          // cells covered by a merge: hMerge / vMerge
          let merge = '';
          if (cell.merged) {
            let hm = false;
            let vm = false;
            for (let rr = 0; rr <= r; rr++)
              for (let cc = 0; cc <= c; cc++) {
                const a = e.rows[rr].cells[cc];
                if (!a.merged && (a.rowSpan ?? 1) + rr > r && (a.colSpan ?? 1) + cc > c) {
                  if (rr === r) hm = true;
                  else if (cc === c) vm = true;
                  else vm = hm = true;
                }
              }
            merge = `${hm ? ' hMerge="1"' : ''}${vm ? ' vMerge="1"' : ''}`;
          }
          const span = `${(cell.rowSpan ?? 1) > 1 ? ` rowSpan="${cell.rowSpan}"` : ''}${(cell.colSpan ?? 1) > 1 ? ` gridSpan="${cell.colSpan}"` : ''}`;
          const extra: Omit<Run, 'text'> = {};
          if (look.color) extra.color = look.color;
          if (look.bold) extra.b = true;
          const body: TextBody = cell.text;
          const txt = `<a:txBody><a:bodyPr/><a:lstStyle/>${parasXml(body, tctx, ctx, { ...extra, ...(body.defaults ?? {}) })}</a:txBody>`;
          const borders = cell.borders ?? look.borders;
          const fill = cell.fill ?? look.fill;
          const anchor = body.anchor === 'm' ? ' anchor="ctr"' : body.anchor === 'b' ? ' anchor="b"' : '';
          const tcPr = `<a:tcPr${anchor}>${cellBorder('lnL', borders[3])}${cellBorder('lnR', borders[1])}${cellBorder('lnT', borders[0])}${cellBorder('lnB', borders[2])}${fillXml(fill ?? { type: 'none' })}</a:tcPr>`;
          return `<a:tc${span}${merge}>${txt}${tcPr}</a:tc>`;
        })
        .join('');
      return `<a:tr h="${emu(row.h)}">${cells}</a:tr>`;
    })
    .join('');
  const grid = `<a:tblGrid>${e.cols.map((w) => `<a:gridCol w="${emu(w)}"/>`).join('')}</a:tblGrid>`;
  return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(e, ctx)}<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>${phXml(e)}</p:nvGraphicFramePr>${xfrm(e, 'p:xfrm')}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr${flags}>${styleId ? `<a:tableStyleId>${styleId}</a:tableStyleId>` : ''}</a:tblPr>${grid}${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

async function chartXmlPart(e: ChartEl, ctx: SlideCtx): Promise<string> {
  const n = ++ctx.charts.n;
  const chartPath = `ppt/charts/chart${n}.xml`;
  const rels = new Rels();
  let ext: string | undefined;
  try {
    const wb = await chartWorkbook(e.chart);
    const wbPath = `ppt/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
    ctx.parts.push({ path: wbPath, data: wb });
    ext = rels.add('package', `../embeddings/Microsoft_Excel_Worksheet${n}.xlsx`);
    ctx.ct.add('xlsx');
  } catch {
    ext = undefined;
  }
  ctx.parts.push({ path: chartPath, data: chartSpaceXml(e.chart, colorXml, ext) });
  if (rels.items.length) ctx.parts.push({ path: `ppt/charts/_rels/chart${n}.xml.rels`, data: rels.xml() });
  const rid = ctx.rels.add('chart', `../charts/chart${n}.xml`);
  return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(e, ctx)}<p:cNvGraphicFramePr/>${phXml(e)}</p:nvGraphicFramePr>${xfrm(e, 'p:xfrm')}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

async function groupXml(e: GroupEl, ctx: SlideCtx): Promise<string> {
  const kids = (await Promise.all(e.children.map((c) => elXml(c, ctx)))).join('');
  return `<p:grpSp><p:nvGrpSpPr>${cNvPr(e, ctx)}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrm(e, 'a:xfrm', true)}</p:grpSpPr>${kids}</p:grpSp>`;
}

async function elXml(e: El, ctx: SlideCtx): Promise<string> {
  switch (e.type) {
    case 'shape':
      return shapeXml(e, ctx);
    case 'image':
      return pictureXml(e, ctx);
    case 'table':
      return tableXml(e, ctx);
    case 'chart':
      return chartXmlPart(e, ctx);
    case 'group':
      return groupXml(e, ctx);
  }
}

function spTree(content: string): string {
  return `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${content}</p:spTree>`;
}

async function bgXml(f: Fill | undefined, ctx: SlideCtx): Promise<string> {
  if (!f) return '';
  let blip: string | undefined;
  if (f.type === 'image') {
    const p = await ctx.media.add(f.src);
    if (p) blip = ctx.rels.add('image', `../media/${p.split('/').pop()}`);
  }
  return `<p:bg><p:bgPr>${fillXml(f, blip)}<a:effectLst/></p:bgPr></p:bg>`;
}

/* ============================================================ transitions */

function transitionXml(t: Transition | undefined): string {
  if (!t || (t.type === 'none' && !t.after)) return '';
  const dur = t.dur ?? 700;
  const spd = dur < 500 ? 'fast' : dur < 1000 ? 'med' : 'slow';
  const adv = `${t.onClick === false ? ' advClick="0"' : ''}${t.after ? ` advTm="${Math.round(t.after)}"` : ''}`;
  const dir = t.dir ?? 'l';
  let el = '';
  let ext = '';
  switch (t.type) {
    case 'fade':
      el = '<p:fade/>';
      break;
    case 'push':
      el = `<p:push dir="${dir}"/>`;
      break;
    case 'wipe':
      el = `<p:wipe dir="${dir}"/>`;
      break;
    case 'split':
      el = `<p:split orient="${dir === 'u' || dir === 'd' ? 'horz' : 'vert'}" dir="out"/>`;
      break;
    case 'cover':
      el = `<p:cover dir="${dir}"/>`;
      break;
    case 'reveal':
      el = `<p:pull dir="${dir}"/>`;
      break;
    case 'zoom':
      el = '<p:zoom/>';
      break;
    case 'circle':
      el = '<p:circle/>';
      break;
    case 'dissolve':
      el = '<p:dissolve/>';
      break;
    case 'flip':
      ext = `<p14:flip dir="${dir === 'r' ? 'r' : 'l'}"/>`;
      el = '<p:fade/>';
      break;
    case 'morph':
      return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159"><p:transition spd="slow" p14:dur="${dur}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"${adv}><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="${spd}"${adv}><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`;
    default:
      el = '';
  }
  return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition spd="${spd}" p14:dur="${dur}"${adv}>${ext || el}</p:transition></mc:Choice><mc:Fallback><p:transition spd="${spd}"${adv}>${el}</p:transition></mc:Fallback></mc:AlternateContent>`;
}

/* ============================================================= animations */

const SUBTYPE: Record<string, number> = { u: 1, r: 2, d: 4, l: 8 };

function behaviours(a: Anim, spid: number, ctx: { id: () => number }, para?: number): { presetID: number; subtype: number; xml: string } {
  const tgt = `<p:tgtEl><p:spTgt spid="${spid}">${para !== undefined ? `<p:txEl><p:pRg st="${para}" end="${para}"/></p:txEl>` : ''}</p:spTgt></p:tgtEl>`;
  const dur = Math.max(1, a.dur);
  const vis = (v: 'visible' | 'hidden', delay = 0) => `<p:set><p:cBhvr><p:cTn id="${ctx.id()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${v}"/></p:to></p:set>`;
  const fx = (transition: 'in' | 'out', filter: string) => `<p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr><p:cTn id="${ctx.id()}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>`;
  const prop = (attr: string, from: string, to: string) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${ctx.id()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const dir = a.dir ?? 'd';
  const offX = dir === 'l' ? '0-#ppt_w/2' : dir === 'r' ? '1+#ppt_w/2' : '#ppt_x';
  const offY = dir === 'd' ? '1+#ppt_h/2' : dir === 'u' ? '0-#ppt_h/2' : '#ppt_y';
  const wipeDir = { d: 'up', u: 'down', l: 'right', r: 'left' }[dir];
  const entr = a.cls === 'entr';
  switch (a.effect) {
    case 'appear':
      return { presetID: 1, subtype: 0, xml: vis('visible') };
    case 'disappear':
      return { presetID: 1, subtype: 0, xml: vis('hidden') };
    case 'fade':
      return entr ? { presetID: 10, subtype: 0, xml: vis('visible') + fx('in', 'fade') } : { presetID: 10, subtype: 0, xml: fx('out', 'fade') + vis('hidden', dur - 1) };
    case 'fly':
      return entr
        ? { presetID: 2, subtype: SUBTYPE[dir], xml: vis('visible') + prop('ppt_x', offX, '#ppt_x') + prop('ppt_y', offY, '#ppt_y') }
        : { presetID: 2, subtype: SUBTYPE[dir], xml: prop('ppt_x', '#ppt_x', offX) + prop('ppt_y', '#ppt_y', offY) + vis('hidden', dur - 1) };
    case 'float':
      return { presetID: 42, subtype: 0, xml: vis('visible') + fx('in', 'fade') + prop('ppt_x', '#ppt_x', '#ppt_x') + prop('ppt_y', `#ppt_y${a.dir === 'd' ? '-' : '+'}.1`, '#ppt_y') };
    case 'zoom':
      return entr
        ? { presetID: 53, subtype: 16, xml: vis('visible') + prop('ppt_w', '0', '#ppt_w') + prop('ppt_h', '0', '#ppt_h') + fx('in', 'fade') }
        : { presetID: 53, subtype: 32, xml: prop('ppt_w', '#ppt_w', '0') + prop('ppt_h', '#ppt_h', '0') + fx('out', 'fade') + vis('hidden', dur - 1) };
    case 'wipe':
      return entr ? { presetID: 22, subtype: SUBTYPE[dir], xml: vis('visible') + fx('in', `wipe(${wipeDir})`) } : { presetID: 22, subtype: SUBTYPE[dir], xml: fx('out', `wipe(${wipeDir})`) + vis('hidden', dur - 1) };
    case 'split':
      return { presetID: 16, subtype: 21, xml: vis('visible') + fx('in', 'barn(outVertical)') };
    case 'wheel':
      return { presetID: 21, subtype: 1, xml: vis('visible') + fx('in', 'wheel(1)') };
    case 'bounce':
      return { presetID: 26, subtype: 0, xml: vis('visible') + fx('in', 'fade') + prop('ppt_y', '#ppt_y-0.25', '#ppt_y') };
    case 'pulse':
      return { presetID: 26, subtype: 0, xml: `<p:animScale><p:cBhvr><p:cTn id="${ctx.id()}" dur="${Math.round(dur / 2)}" autoRev="1" fill="hold"/>${tgt}</p:cBhvr><p:by x="105000" y="105000"/></p:animScale>` };
    case 'spin':
      return { presetID: 8, subtype: 0, xml: `<p:animRot by="21600000"><p:cBhvr><p:cTn id="${ctx.id()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>` };
    case 'grow':
      return { presetID: 6, subtype: 0, xml: `<p:animScale><p:cBhvr><p:cTn id="${ctx.id()}" dur="${Math.round(dur / 2)}" autoRev="1" fill="hold"/>${tgt}</p:cBhvr><p:by x="135000" y="135000"/></p:animScale>` };
    case 'teeter':
      return { presetID: 32, subtype: 0, xml: `<p:animRot by="300000"><p:cBhvr><p:cTn id="${ctx.id()}" dur="${Math.round(dur / 4)}" autoRev="1" repeatCount="2000" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>` };
    case 'transparency':
      return { presetID: 9, subtype: 0, xml: `<p:set><p:cBhvr><p:cTn id="${ctx.id()}" dur="${dur}" autoRev="1" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="0.5"/></p:to></p:set>` };
    default:
      return { presetID: 10, subtype: 0, xml: vis('visible') + fx('in', 'fade') };
  }
}

function timingXml(slide: Slide, ids: Map<string, number>): string {
  if (!slide.anims?.length) return '';
  const build = buildOf(slide);
  let n = 3;
  const id = () => n++;
  const bld = new Map<number, boolean>();
  const effectNode = (it: (typeof build.steps)[number][number], nodeType: string, grp: number) => {
    const spid = ids.get(it.anim.el);
    if (!spid) return '';
    const eid = id();
    const b = behaviours(it.anim, spid, { id }, it.para);
    if (it.para !== undefined) bld.set(spid, true);
    else if (!bld.has(spid)) bld.set(spid, false);
    return `<p:par><p:cTn id="${eid}" presetID="${b.presetID}" presetClass="${it.anim.cls}" presetSubtype="${b.subtype}" fill="hold"${it.para !== undefined ? '' : ` grpId="${grp}"`} nodeType="${nodeType}"><p:stCondLst><p:cond delay="${it.anim.delay}"/></p:stCondLst><p:childTnLst>${b.xml}</p:childTnLst></p:cTn></p:par>`;
  };
  const stepXml = (items: (typeof build.steps)[number], click: boolean, first: boolean) => {
    // group items by start time: "with" items share their group, "after" items start new groups
    const groups: { at: number; items: typeof items }[] = [];
    for (const it of items) {
      const start = it.offset - it.anim.delay;
      const g = groups.find((x) => Math.abs(x.at - start) < 1);
      if (g) g.items.push(it);
      else groups.push({ at: start, items: [it] });
    }
    const outer = id();
    const cond = click ? '<p:cond delay="indefinite"/>' : first ? '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' : '<p:cond delay="0"/>';
    const inner = groups
      .map((g, gi) => {
        const gid = id();
        const nodes = g.items.map((it, ii) => effectNode(it, gi === 0 && ii === 0 ? (click ? 'clickEffect' : it.anim.start === 'with' ? 'withEffect' : 'afterEffect') : ii === 0 ? 'afterEffect' : 'withEffect', 0)).join('');
        return `<p:par><p:cTn id="${gid}" fill="hold"><p:stCondLst><p:cond delay="${Math.round(g.at)}"/></p:stCondLst><p:childTnLst>${nodes}</p:childTnLst></p:cTn></p:par>`;
      })
      .join('');
    return `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
  };
  const steps: string[] = [];
  if (build.auto) steps.push(stepXml(build.auto, false, true));
  build.steps.forEach((s) => steps.push(stepXml(s, true, false)));
  const bldLst = [...bld.entries()].map(([spid, byPara]) => (byPara ? `<p:bldP spid="${spid}" grpId="0" build="p"/>` : `<p:bldP spid="${spid}" grpId="0" animBg="1"/>`)).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${steps.join('')}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>${bldLst ? `<p:bldLst>${bldLst}</p:bldLst>` : ''}</p:timing>`;
}

/* ================================================================ masters */

function themeXml(pres: Presentation): string {
  const c = pres.theme.colors;
  const clr = (tag: string, v: string) => `<a:${tag}><a:srgbClr val="${v.replace('#', '').slice(0, 6).toUpperCase()}"/></a:${tag}>`;
  const font = (tag: string, f: string) => `<a:${tag}><a:latin typeface="${x(f)}"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`;
  return `${HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${x(pres.theme.name)}"><a:themeElements><a:clrScheme name="${x(pres.theme.name)}">${clr('dk1', c.tx1)}${clr('lt1', c.bg1)}${clr('dk2', c.tx2)}${clr('lt2', c.bg2)}${clr('accent1', c.accent1)}${clr('accent2', c.accent2)}${clr('accent3', c.accent3)}${clr('accent4', c.accent4)}${clr('accent5', c.accent5)}${clr('accent6', c.accent6)}${clr('hlink', c.hlink)}${clr('folHlink', c.folHlink)}</a:clrScheme><a:fontScheme name="${x(pres.theme.name)}">${font('majorFont', pres.theme.fonts.major)}${font('minorFont', pres.theme.fonts.minor)}</a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

function levelStyles(kind: 'title' | 'body' | 'other', pres: Presentation): string {
  const t = pres.master.text;
  const out: string[] = [];
  for (let i = 0; i < 9; i++) {
    if (kind === 'title') {
      if (i > 0) break;
      out.push(`<a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="${Math.round(t.title.size * 100)}" kern="1200"${t.title.bold ? ' b="1"' : ''}><a:solidFill>${colorXml(t.title.color)}</a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>`);
    } else if (kind === 'body') {
      const sizes = [28, 24, 20, 20, 20, 20, 20, 20, 20].map((s) => Math.round((s * t.body.size) / 28));
      const bullets = ['•', '–', '•', '–', '»', '•', '–', '•', '–'];
      out.push(`<a:lvl${i + 1}pPr marL="${228600 + i * 457200}" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="${Math.round((t.body.lineSpacing ?? 0.9) * 100000)}"/></a:lnSpc><a:spcBef><a:spcPts val="${i === 0 ? 1000 : 500}"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="${bullets[i]}"/><a:defRPr sz="${sizes[i] * 100}" kern="1200"><a:solidFill>${colorXml(t.body.color)}</a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`);
    } else {
      out.push(`<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="${Math.round(t.other.size * 100)}" kern="1200"><a:solidFill>${colorXml(t.other.color)}</a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`);
    }
  }
  return out.join('');
}

const W_EMU = (pres: Presentation) => emu(pres.size.w);
const H_EMU = (pres: Presentation) => emu(pres.size.h);

function masterPlaceholders(pres: Presentation): string {
  const sx = pres.size.w / 1280;
  const sy = pres.size.h / 720;
  const ph = (id: number, name: string, type: string, idx: string, xx: number, y: number, w: number, h: number, body: string) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}"${idx}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(xx * sx)}" y="${emu(y * sy)}"/><a:ext cx="${emu(w * sx)}" cy="${emu(h * sy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${body}</p:txBody></p:sp>`;
  const footerBody = (algn: string, field?: string) => `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="${algn}"><a:defRPr sz="1200"><a:solidFill>${colorXml('@tx1+45')}</a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p>${field ? `<a:fld id="{${guid(field)}}" type="${field}"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>` : ''}<a:endParaRPr lang="en-US"/></a:p>`;
  return (
    ph(2, 'Title Placeholder 1', 'title', '', 88, 38, 1104, 139, '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master title style</a:t></a:r></a:p>') +
    ph(3, 'Text Placeholder 2', 'body', ' idx="1"', 88, 192, 1104, 457, '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>Second level</a:t></a:r></a:p>') +
    ph(4, 'Date Placeholder 3', 'dt', ' sz="half" idx="10"', 88, 667, 288, 38, footerBody('l', 'datetime1')) +
    ph(5, 'Footer Placeholder 4', 'ftr', ' sz="quarter" idx="11"', 424, 667, 432, 38, footerBody('ctr')) +
    ph(6, 'Slide Number Placeholder 5', 'sldNum', ' sz="quarter" idx="12"', 904, 667, 288, 38, footerBody('r', 'slidenum'))
  );
}

async function masterXml(pres: Presentation, layoutRids: string[], ctx: SlideCtx): Promise<string> {
  const decor = (await Promise.all(pres.master.decor.map((d) => elXml(d, ctx)))).join('');
  const bg = await bgXml(pres.master.background, ctx);
  return `${HEAD}<p:sldMaster ${NS}><p:cSld>${bg}${spTree(decor + masterPlaceholders(pres))}</p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst>${layoutRids.map((rid, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="${rid}"/>`).join('')}</p:sldLayoutIdLst><p:txStyles><p:titleStyle>${levelStyles('title', pres)}</p:titleStyle><p:bodyStyle>${levelStyles('body', pres)}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${levelStyles('other', pres)}</p:otherStyle></p:txStyles></p:sldMaster>`;
}

/** Layout placeholder text styles (the effective defaults of each placeholder). */
function placeholderLst(e: El, pres: Presentation): string {
  if (e.type !== 'shape' || !e.text) return '<a:lstStyle/>';
  const tctx: TextContext = { pres, role: roleOf(e.ph), ph: e.ph };
  const p0 = e.text.paras[0] ?? { runs: [] };
  const r = effectiveRun({}, p0, e.text, tctx);
  const at: string[] = [`sz="${Math.round(r.size * 100)}"`];
  if (r.b) at.push('b="1"');
  const pe = effectivePara(p0, tctx);
  const noBullet = pe.bullet?.type === 'none';
  const algn = p0.align ? ` algn="${p0.align === 'center' ? 'ctr' : p0.align === 'right' ? 'r' : p0.align === 'justify' ? 'just' : 'l'}"` : '';
  const bullets = noBullet ? '<a:buNone/>' : '';
  const mar = noBullet ? ` marL="0" indent="0"` : '';
  // subtitles have no space before their first paragraph, unlike the master's body style
  const spacing = tctx.role === 'body' ? `<a:spcBef><a:spcPts val="${Math.round(pe.spaceBefore * 100)}"/></a:spcBef>` : '';
  const color = e.text.defaults?.color ? `<a:solidFill>${colorXml(e.text.defaults.color)}</a:solidFill>` : '';
  return `<a:lstStyle><a:lvl1pPr${mar}${algn}>${spacing}${bullets}<a:defRPr ${at.join(' ')}>${color}</a:defRPr></a:lvl1pPr></a:lstStyle>`;
}

async function layoutXml(pres: Presentation, l: Layout, ctx: SlideCtx): Promise<string> {
  const decor = (await Promise.all(l.decor.map((d) => elXml(d, ctx)))).join('');
  const phs = l.placeholders
    .map((p) => {
      const id = ctx.nextId++;
      const body = p.type === 'shape' && p.text ? p.text : undefined;
      const prompt = p.ph === 'title' || p.ph === 'ctrTitle' ? 'Click to edit Master title style' : p.ph === 'subTitle' ? 'Click to edit Master subtitle style' : 'Click to edit Master text styles';
      return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${x(p.name ?? 'Placeholder')}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>${phXml(p)}</p:nvSpPr><p:spPr>${xfrm(p)}</p:spPr><p:txBody>${bodyPrXml(body, false)}${placeholderLst(p, pres)}<a:p><a:r><a:rPr lang="en-US"/><a:t>${prompt}</a:t></a:r></a:p></p:txBody></p:sp>`;
    })
    .join('');
  const bg = l.background ? await bgXml(l.background, ctx) : '';
  return `${HEAD}<p:sldLayout ${NS} type="${l.type}" preserve="1"><p:cSld name="${x(l.name)}">${bg}${spTree(decor + phs)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function footerShapes(pres: Presentation, slide: Slide, index: number, ctx: SlideCtx): string {
  const f = pres.footer;
  if (!f || (!f.text && !f.slideNumber && !f.date)) return '';
  const layout = pres.layouts.find((l) => l.id === slide.layout);
  if (f.skipTitle && layout?.type === 'title') return '';
  const sx = pres.size.w / 1280;
  const y = pres.size.h - 53;
  const sp = (type: string, idx: number, xx: number, w: number, content: string) => {
    const id = ctx.nextId++;
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${type} ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}" sz="quarter" idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(xx * sx)}" y="${emu(y)}"/><a:ext cx="${emu(w * sx)}" cy="${emu(38)}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${content}<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
  };
  let out = '';
  if (f.date) out += sp('dt', 10, 88, 288, `<a:fld id="{${guid(`dt${index}`)}}" type="datetime1"><a:rPr lang="en-US"/><a:t>${x(new Date().toLocaleDateString())}</a:t></a:fld>`);
  if (f.text) out += sp('ftr', 11, 424, 432, `<a:r><a:rPr lang="en-US"/><a:t>${x(f.text)}</a:t></a:r>`);
  if (f.slideNumber) out += sp('sldNum', 12, 904, 288, `<a:fld id="{${guid(`sn${index}`)}}" type="slidenum"><a:rPr lang="en-US"/><a:t>${index + 1}</a:t></a:fld>`);
  return out;
}

function notesMasterXml(): string {
  return `${HEAD}<p:notesMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${spTree(
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>',
  )}</p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
}

function notesSlideXml(notes: string): string {
  const paras = notes
    .split('\n')
    .map((l) => (l ? `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${x(l)}</a:t></a:r></a:p>` : '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>'))
    .join('');
  return `${HEAD}<p:notes ${NS}><p:cSld>${spTree(
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>`,
  )}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

/* ================================================================= export */

export async function exportPptx(pres: Presentation): Promise<Uint8Array> {
  const parts: Part[] = [];
  const media = new Media();
  const ct = new Set<string>();
  const charts = { n: 0 };
  const mk = (rels: Rels, slideIndex: number): SlideCtx => ({ rels, media, pres, slideIndex, ids: new Map(), nextId: 2, slidePath: (i) => `../slides/slide${i + 1}.xml`, charts, parts, ct });

  // layouts
  const layouts = pres.layouts.length ? pres.layouts : [{ id: 'blank', name: 'Blank', type: 'blank' as const, decor: [], placeholders: [] }];
  const layoutIndex = new Map(layouts.map((l, i) => [l.id, i]));
  const masterRels = new Rels();
  const layoutRids: string[] = [];
  for (let i = 0; i < layouts.length; i++) {
    const rels = new Rels();
    rels.add('slideMaster', '../slideMasters/slideMaster1.xml');
    const ctx = mk(rels, -1);
    parts.push({ path: `ppt/slideLayouts/slideLayout${i + 1}.xml`, data: await layoutXml(pres, layouts[i], ctx) });
    parts.push({ path: `ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, data: rels.xml() });
    layoutRids.push(masterRels.add('slideLayout', `../slideLayouts/slideLayout${i + 1}.xml`));
  }
  masterRels.add('theme', '../theme/theme1.xml');
  const mctx = mk(masterRels, -1);
  mctx.nextId = 100;
  parts.push({ path: 'ppt/slideMasters/slideMaster1.xml', data: await masterXml(pres, layoutRids, mctx) });
  parts.push({ path: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: masterRels.xml() });
  parts.push({ path: 'ppt/theme/theme1.xml', data: themeXml(pres) });

  // slides
  const presRels = new Rels();
  presRels.add('slideMaster', 'slideMasters/slideMaster1.xml');
  const hasNotes = pres.slides.some((s) => s.notes?.trim());
  const slideRids: string[] = [];
  for (let i = 0; i < pres.slides.length; i++) {
    const s = pres.slides[i];
    const rels = new Rels();
    rels.add('slideLayout', `../slideLayouts/slideLayout${(layoutIndex.get(s.layout) ?? 0) + 1}.xml`);
    const ctx = mk(rels, i);
    const els = (await Promise.all(s.elements.map((e) => elXml(e, ctx)))).join('') + footerShapes(pres, s, i, ctx);
    const bg = await bgXml(s.background, ctx);
    const attrs = `${s.hidden ? ' show="0"' : ''}${s.hideDecor ? ' showMasterSp="0"' : ''}`;
    const xml = `${HEAD}<p:sld ${NS}${attrs}><p:cSld>${bg}${spTree(els)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transitionXml(s.transition)}${timingXml(s, ctx.ids)}</p:sld>`;
    if (s.notes?.trim()) {
      const nrels = new Rels();
      nrels.add('notesMaster', '../notesMasters/notesMaster1.xml');
      nrels.add('slide', `../slides/slide${i + 1}.xml`);
      parts.push({ path: `ppt/notesSlides/notesSlide${i + 1}.xml`, data: notesSlideXml(s.notes) });
      parts.push({ path: `ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`, data: nrels.xml() });
      rels.add('notesSlide', `../notesSlides/notesSlide${i + 1}.xml`);
    }
    parts.push({ path: `ppt/slides/slide${i + 1}.xml`, data: xml });
    parts.push({ path: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: rels.xml() });
    slideRids.push(presRels.add('slide', `slides/slide${i + 1}.xml`));
  }
  let notesRid = '';
  if (hasNotes) {
    const nm = new Rels();
    nm.add('theme', '../theme/theme2.xml');
    parts.push({ path: 'ppt/notesMasters/notesMaster1.xml', data: notesMasterXml() });
    parts.push({ path: 'ppt/notesMasters/_rels/notesMaster1.xml.rels', data: nm.xml() });
    parts.push({ path: 'ppt/theme/theme2.xml', data: themeXml(pres) });
    notesRid = presRels.add('notesMaster', 'notesMasters/notesMaster1.xml');
  }
  presRels.add('theme', 'theme/theme1.xml');
  presRels.add('presProps', 'presProps.xml');
  presRels.add('viewProps', 'viewProps.xml');
  presRels.add('tableStyles', 'tableStyles.xml');

  const sizeType = Math.abs(pres.size.w / pres.size.h - 16 / 9) < 0.01 ? '' : Math.abs(pres.size.w / pres.size.h - 4 / 3) < 0.01 ? ' type="screen4x3"' : ' type="custom"';
  const presentation = `${HEAD}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>${notesRid ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesRid}"/></p:notesMasterIdLst>` : ''}<p:sldIdLst>${slideRids.map((rid, i) => `<p:sldId id="${256 + i}" r:id="${rid}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${W_EMU(pres)}" cy="${H_EMU(pres)}"${sizeType}/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${levelStyles('other', pres)}</p:defaultTextStyle></p:presentation>`;
  parts.push({ path: 'ppt/presentation.xml', data: presentation });
  parts.push({ path: 'ppt/_rels/presentation.xml.rels', data: presRels.xml() });
  parts.push({ path: 'ppt/presProps.xml', data: `${HEAD}<p:presentationPr ${NS}/>` });
  parts.push({ path: 'ppt/viewProps.xml', data: `${HEAD}<p:viewPr ${NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>` });
  parts.push({ path: 'ppt/tableStyles.xml', data: `${HEAD}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>` });

  // package metadata
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  parts.push({
    path: 'docProps/core.xml',
    data: `${HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${x(pres.props.title ?? '')}</dc:title><dc:creator>${x(pres.props.author ?? '')}</dc:creator><cp:lastModifiedBy>Affice</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${x(pres.props.created ?? now)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  });
  parts.push({ path: 'docProps/app.xml', data: `${HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Affice</Application><Slides>${pres.slides.length}</Slides><PresentationFormat>On-screen Show</PresentationFormat></Properties>` });
  parts.push({
    path: '_rels/.rels',
    data: `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  });

  // content types
  for (const f of media.files) ct.add(f.ext);
  const mime: Record<string, string> = { png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', tiff: 'image/tiff', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
  const over = (path: string, type: string) => `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}"/>`;
  const overrides = parts
    .map((p) => {
      if (p.path === 'ppt/presentation.xml') return over(p.path, 'presentationml.presentation.main+xml');
      if (/^ppt\/slides\/slide\d+\.xml$/.test(p.path)) return over(p.path, 'presentationml.slide+xml');
      if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p.path)) return over(p.path, 'presentationml.slideLayout+xml');
      if (/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p.path)) return over(p.path, 'presentationml.slideMaster+xml');
      if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(p.path)) return over(p.path, 'presentationml.notesSlide+xml');
      if (/^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(p.path)) return over(p.path, 'presentationml.notesMaster+xml');
      if (/^ppt\/theme\/theme\d+\.xml$/.test(p.path)) return over(p.path, 'theme+xml');
      if (/^ppt\/charts\/chart\d+\.xml$/.test(p.path)) return over(p.path, 'drawingml.chart+xml');
      if (p.path === 'ppt/presProps.xml') return over(p.path, 'presentationml.presProps+xml');
      if (p.path === 'ppt/viewProps.xml') return over(p.path, 'presentationml.viewProps+xml');
      if (p.path === 'ppt/tableStyles.xml') return over(p.path, 'presentationml.tableStyles+xml');
      if (p.path === 'docProps/app.xml') return over(p.path, 'extended-properties+xml');
      if (p.path === 'docProps/core.xml') return `<Override PartName="/${p.path}" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`;
      return '';
    })
    .join('');
  const defaults = [...ct].map((e) => `<Default Extension="${e}" ContentType="${mime[e] ?? 'application/octet-stream'}"/>`).join('');
  parts.push({
    path: '[Content_Types].xml',
    data: `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${defaults}${overrides}</Types>`,
  });

  const files: Record<string, Uint8Array> = {};
  for (const p of parts) files[p.path] = typeof p.data === 'string' ? strToU8(p.data) : p.data;
  for (const m of media.files) files[m.path] = m.data;
  return zipSync(files, { level: 6 });
}

