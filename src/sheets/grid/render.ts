import { cssFontStack } from '@/lib/fonts';
import { isErr, asErr, type Scalar } from '../engine/values';
import { formatValue } from '../format/numfmt';
import { cellKey, colName, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import type { Border, CellStyle } from '../model/types';
import type { Sheet } from '../model/workbook';
import type { CellPos, SheetDoc, SheetSelection } from '../doc';
import type { CfEffect, CfEvaluator } from './cf';
import type { SheetGeometry } from './geometry';

export interface GridTheme {
  bg: string;
  line: string;
  text: string;
  headerBg: string;
  headerFg: string;
  headerLine: string;
  headerSelBg: string;
  headerSelFg: string;
  accent: string;
  selFill: string;
  frozenLine: string;
  uiFont: string;
  dark: boolean;
}

export interface Viewport {
  width: number;
  height: number;
  headerW: number;
  headerH: number;
  zoom: number;
  scrollX: number;
  scrollY: number;
  fr: number;
  fc: number;
}

export interface Overlay {
  sel: SheetSelection;
  editing?: CellPos | null;
  refs?: { range: Range; color: string }[];
  clip?: { range: Range; phase: number } | null;
  fillPreview?: Range | null;
  movePreview?: Range | null;
  showFormulas?: boolean;
  focused: boolean;
  findHits?: Set<number>;
  spillHint?: Range | null;
}

/* =============================================================== mapping */

export class Mapper {
  readonly fcW: number;
  readonly frH: number;
  readonly bodyX0: number;
  readonly bodyY0: number;
  constructor(
    public geo: SheetGeometry,
    public vp: Viewport,
  ) {
    this.fcW = geo.cols.pos(vp.fc);
    this.frH = geo.rows.pos(vp.fr);
    this.bodyX0 = vp.headerW + this.fcW * vp.zoom;
    this.bodyY0 = vp.headerH + this.frH * vp.zoom;
  }
  x(c: number): number {
    const z = this.vp.zoom;
    if (c < this.vp.fc) return this.vp.headerW + this.geo.cols.pos(c) * z;
    return this.bodyX0 + (this.geo.cols.pos(c) - this.fcW - this.vp.scrollX) * z;
  }
  y(r: number): number {
    const z = this.vp.zoom;
    if (r < this.vp.fr) return this.vp.headerH + this.geo.rows.pos(r) * z;
    return this.bodyY0 + (this.geo.rows.pos(r) - this.frH - this.vp.scrollY) * z;
  }
  /** Cell under a screen point (in grid-local css px). */
  cellAt(px: number, py: number): CellPos {
    return { r: this.rowAt(py), c: this.colAt(px) };
  }
  colAt(px: number): number {
    const z = this.vp.zoom;
    if (px < this.bodyX0 && this.vp.fc > 0) return this.geo.cols.at(Math.max(0, (px - this.vp.headerW) / z));
    return this.geo.cols.at(this.fcW + this.vp.scrollX + Math.max(0, px - this.bodyX0) / z);
  }
  rowAt(py: number): number {
    const z = this.vp.zoom;
    if (py < this.bodyY0 && this.vp.fr > 0) return this.geo.rows.at(Math.max(0, (py - this.vp.headerH) / z));
    return this.geo.rows.at(this.frH + this.vp.scrollY + Math.max(0, py - this.bodyY0) / z);
  }
  /** First and last visible body rows/cols. */
  bodyRows(): [number, number] {
    const z = this.vp.zoom;
    const top = this.frH + this.vp.scrollY;
    const r0 = Math.max(this.vp.fr, this.geo.rows.at(top));
    const r1 = this.geo.rows.at(top + Math.max(0, this.vp.height - this.bodyY0) / z);
    return [r0, Math.min(MAX_ROWS - 1, r1)];
  }
  bodyCols(): [number, number] {
    const z = this.vp.zoom;
    const left = this.fcW + this.vp.scrollX;
    const c0 = Math.max(this.vp.fc, this.geo.cols.at(left));
    const c1 = this.geo.cols.at(left + Math.max(0, this.vp.width - this.bodyX0) / z);
    return [c0, Math.min(MAX_COLS - 1, c1)];
  }
  rect(rg: Range): { x: number; y: number; w: number; h: number } {
    const x = this.x(rg.c1);
    const y = this.y(rg.r1);
    return { x, y, w: this.x(rg.c2 + 1) - x, h: this.y(rg.r2 + 1) - y };
  }
}

/* ============================================================= measuring */

const widths = new Map<string, number>();
const metrics = new Map<string, { ascent: number; descent: number }>();

export function measure(ctx: CanvasRenderingContext2D, font: string, text: string): number {
  const key = font + '\u0000' + text;
  let w = widths.get(key);
  if (w === undefined) {
    if (widths.size > 30000) widths.clear();
    ctx.font = font;
    w = ctx.measureText(text).width;
    widths.set(key, w);
  }
  return w;
}

function fontMetrics(ctx: CanvasRenderingContext2D, font: string, px: number): { ascent: number; descent: number } {
  let m = metrics.get(font);
  if (!m) {
    ctx.font = font;
    const tm = ctx.measureText('Hgjy');
    m = { ascent: tm.fontBoundingBoxAscent || px * 0.8, descent: tm.fontBoundingBoxDescent || px * 0.22 };
    metrics.set(font, m);
  }
  return m;
}

export function clearMeasureCache(): void {
  widths.clear();
  metrics.clear();
}

export const DEFAULT_FONT = 'Calibri';
export const DEFAULT_SIZE = 11;

export function cellFont(st: CellStyle, zoom: number): { font: string; px: number } {
  const px = Math.max(1, ((st.size ?? DEFAULT_SIZE) * 96) / 72) * zoom;
  const font = `${st.italic ? 'italic ' : ''}${st.bold ? 'bold ' : ''}${px.toFixed(2)}px ${cssFontStack(st.font ?? DEFAULT_FONT)}`;
  return { font, px };
}

/** Formats a number to fit a width (General shrinks decimals; other formats show ####). */
export function fitText(ctx: CanvasRenderingContext2D, font: string, v: number, fmt: string | undefined, avail: number, text: string): string {
  if (measure(ctx, font, text) <= avail) return text;
  const general = !fmt || /^general$/i.test(fmt);
  if (general && Number.isFinite(v)) {
    const abs = Math.abs(v);
    if (!Number.isInteger(v) || abs >= 1e11) {
      for (let p = 10; p >= 1; p--) {
        let t: string;
        if (abs >= 1e11 || abs < 1e-5) t = v.toExponential(Math.max(0, p - 1)).replace(/e\+?(-?)(\d)$/, 'E+$1' + '0$2').replace('E+-', 'E-').toUpperCase();
        else t = String(parseFloat(v.toPrecision(p)));
        if (measure(ctx, font, t) <= avail) return t;
      }
    }
  }
  const hash = measure(ctx, font, '#') || 7;
  return '#'.repeat(Math.max(1, Math.floor(avail / hash)));
}

function wrapLines(ctx: CanvasRenderingContext2D, font: string, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      const cand = line + w;
      if (line && measure(ctx, font, cand.trimEnd()) > width) {
        out.push(line.trimEnd());
        line = w.trimStart();
        // break very long words
        while (measure(ctx, font, line) > width && line.length > 1) {
          let k = line.length - 1;
          while (k > 1 && measure(ctx, font, line.slice(0, k)) > width) k--;
          out.push(line.slice(0, k));
          line = line.slice(k);
        }
      } else line = cand;
    }
    out.push(line.trimEnd());
  }
  return out;
}

/* ============================================================== borders */

const BORDER_W: Record<string, number> = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };

function heavier(a: Border | undefined, b: Border | undefined): Border | undefined {
  if (!a) return b;
  if (!b) return a;
  return (BORDER_W[b.style] ?? 1) > (BORDER_W[a.style] ?? 1) ? b : a;
}

function strokeBorder(ctx: CanvasRenderingContext2D, b: Border, x1: number, y1: number, x2: number, y2: number, dpr: number, color: string) {
  ctx.strokeStyle = color;
  const dash: Record<string, number[]> = {
    dotted: [1, 1],
    hair: [1, 1],
    dashed: [3, 2],
    mediumDashed: [4, 2],
    dashDot: [4, 2, 1, 2],
    mediumDashDot: [5, 2, 2, 2],
    dashDotDot: [4, 2, 1, 2, 1, 2],
    mediumDashDotDot: [5, 2, 2, 2, 2, 2],
    slantDashDot: [5, 2, 2, 2],
  };
  ctx.setLineDash(dash[b.style] ?? []);
  const snap = (v: number) => Math.round(v * dpr) / dpr;
  const horiz = y1 === y2;
  if (b.style === 'double') {
    ctx.lineWidth = 1 / dpr;
    const off = 1;
    ctx.beginPath();
    if (horiz) {
      ctx.moveTo(x1, snap(y1) - off + 0.5 / dpr);
      ctx.lineTo(x2, snap(y1) - off + 0.5 / dpr);
      ctx.moveTo(x1, snap(y1) + off + 0.5 / dpr);
      ctx.lineTo(x2, snap(y1) + off + 0.5 / dpr);
    } else {
      ctx.moveTo(snap(x1) - off + 0.5 / dpr, y1);
      ctx.lineTo(snap(x1) - off + 0.5 / dpr, y2);
      ctx.moveTo(snap(x1) + off + 0.5 / dpr, y1);
      ctx.lineTo(snap(x1) + off + 0.5 / dpr, y2);
    }
    ctx.stroke();
    ctx.setLineDash([]);
    return;
  }
  const cssW = BORDER_W[b.style] ?? 1;
  ctx.lineWidth = cssW;
  const off = (Math.round(cssW * dpr) % 2 ? 0.5 : 0) / dpr;
  ctx.beginPath();
  if (horiz) {
    ctx.moveTo(x1, snap(y1) + off);
    ctx.lineTo(x2, snap(y1) + off);
  } else {
    ctx.moveTo(snap(x1) + off, y1);
    ctx.lineTo(snap(x1) + off, y2);
  }
  ctx.stroke();
  ctx.setLineDash([]);
}

/* ================================================================ icons */

const ICON_COLORS = { red: '#e5484d', yellow: '#f5b301', green: '#30a46c', gray: '#8b8d98' };

function drawIcon(ctx: CanvasRenderingContext2D, set: string, index: number, total: number, x: number, y: number, s: number) {
  const level = total - 1 - index; // 0 = best
  const color = level === 0 ? ICON_COLORS.green : level === total - 1 ? ICON_COLORS.red : ICON_COLORS.yellow;
  ctx.save();
  ctx.translate(x + s / 2, y + s / 2);
  ctx.fillStyle = color;
  ctx.strokeStyle = color;
  if (set.includes('Arrows')) {
    const angle = total === 3 ? [-90, 0, 90][level] : total === 4 ? [-90, -45, 45, 90][level] : [-90, -45, 0, 45, 90][level];
    ctx.rotate((angle * Math.PI) / 180);
    ctx.beginPath();
    ctx.moveTo(s * 0.45, 0);
    ctx.lineTo(0, -s * 0.4);
    ctx.lineTo(0, -s * 0.15);
    ctx.lineTo(-s * 0.45, -s * 0.15);
    ctx.lineTo(-s * 0.45, s * 0.15);
    ctx.lineTo(0, s * 0.15);
    ctx.lineTo(0, s * 0.4);
    ctx.closePath();
    ctx.fill();
  } else if (set.includes('Stars')) {
    ctx.beginPath();
    for (let i = 0; i < 10; i++) {
      const rad = i % 2 ? s * 0.2 : s * 0.45;
      const a = (i * Math.PI) / 5 - Math.PI / 2;
      ctx.lineTo(Math.cos(a) * rad, Math.sin(a) * rad);
    }
    ctx.closePath();
    ctx.fillStyle = level === 0 ? '#f5b301' : level === 1 ? '#f5b30188' : '#c7c7cc';
    ctx.fill();
  } else if (set.includes('Flags')) {
    ctx.fillRect(-s * 0.3, -s * 0.4, s * 0.08, s * 0.8);
    ctx.beginPath();
    ctx.moveTo(-s * 0.22, -s * 0.4);
    ctx.lineTo(s * 0.4, -s * 0.2);
    ctx.lineTo(-s * 0.22, 0);
    ctx.fill();
  } else if (set.includes('Ratings')) {
    for (let i = 0; i < 4; i++) {
      ctx.fillStyle = i <= total - 1 - level - 1 ? '#3e63dd' : '#c7c7cc';
      const bh = (s * 0.2 * (i + 1)) / 1.2;
      ctx.fillRect(-s * 0.45 + i * s * 0.24, s * 0.4 - bh, s * 0.16, bh);
    }
  } else if (set.includes('Symbols')) {
    ctx.beginPath();
    ctx.arc(0, 0, s * 0.42, 0, Math.PI * 2);
    ctx.fill();
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = Math.max(1, s * 0.12);
    ctx.beginPath();
    if (level === 0) {
      ctx.moveTo(-s * 0.2, 0);
      ctx.lineTo(-s * 0.05, s * 0.15);
      ctx.lineTo(s * 0.2, -s * 0.15);
    } else if (level === total - 1) {
      ctx.moveTo(-s * 0.15, -s * 0.15);
      ctx.lineTo(s * 0.15, s * 0.15);
      ctx.moveTo(s * 0.15, -s * 0.15);
      ctx.lineTo(-s * 0.15, s * 0.15);
    } else {
      ctx.moveTo(0, -s * 0.2);
      ctx.lineTo(0, s * 0.05);
      ctx.moveTo(0, s * 0.15);
      ctx.lineTo(0, s * 0.2);
    }
    ctx.stroke();
  } else {
    // traffic lights
    ctx.beginPath();
    ctx.arc(0, 0, s * 0.42, 0, Math.PI * 2);
    ctx.fill();
  }
  ctx.restore();
}

/* ================================================================ render */

export interface RenderInput {
  ctx: CanvasRenderingContext2D;
  doc: SheetDoc;
  sheet: Sheet;
  geo: SheetGeometry;
  vp: Viewport;
  dpr: number;
  theme: GridTheme;
  cf: CfEvaluator;
  ov: Overlay;
}

interface CellInfo {
  st: CellStyle;
  text: string;
  color?: string;
  value: Scalar | undefined;
  cf?: CfEffect;
}

export function renderGrid(inp: RenderInput): void {
  const { ctx, doc, sheet, geo, vp, dpr, theme, cf, ov } = inp;
  const m = new Mapper(geo, vp);
  const z = vp.zoom;
  const W = vp.width;
  const H = vp.height;
  const snap = (v: number) => Math.round(v * dpr) / dpr + 0.5 / dpr;
  ctx.save();
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = theme.bg;
  ctx.fillRect(0, 0, W, H);

  const [br0, br1] = m.bodyRows();
  const [bc0, bc1] = m.bodyCols();
  const rowsSets: [number, number][] = [];
  const colsSets: [number, number][] = [];
  if (vp.fr > 0) rowsSets.push([0, vp.fr - 1]);
  rowsSets.push([br0, br1]);
  if (vp.fc > 0) colsSets.push([0, vp.fc - 1]);
  colsSets.push([bc0, bc1]);

  const tone = theme.dark ? darkTone : (c: string) => c;
  const showFormulas = !!ov.showFormulas || sheet.view.showFormulas;
  const merges = sheet.merges;
  const infoCache = new Map<number, CellInfo | null>();

  const info = (r: number, c: number): CellInfo | null => {
    const k = cellKey(r, c);
    let inf = infoCache.get(k);
    if (inf !== undefined) return inf;
    const sid = doc.wb.cellStyleId(sheet, r, c);
    let st = doc.wb.style(sid);
    const cfx = cf.effect(sheet, r, c);
    if (cfx?.style) st = { ...st, ...cfx.style };
    const cell = sheet.cells.get(k);
    const value = doc.value(sheet, r, c);
    let text = '';
    let color: string | undefined;
    if (showFormulas && cell?.f !== undefined) text = '=' + cell.f;
    else if (value !== undefined && value !== null) {
      if (isErr(value)) text = asErr(value).error;
      else {
        const f = formatValue(value, st.numFmt);
        text = f.text;
        color = f.color;
      }
    }
    if (cfx?.hideValue) text = '';
    if (!text && !st.fill && !cfx?.scale && !cfx?.bar && !cfx?.icon && !st.bt && !st.bb && !st.bl && !st.br && !cell?.note) inf = null;
    else inf = { st, text, color, value, cf: cfx };
    infoCache.set(k, inf);
    return inf;
  };

  // merges visible in the viewport
  const visMerges = merges.filter((mg) => rowsSets.some(([a, b]) => mg.r1 <= b && mg.r2 >= a) && colsSets.some(([a, b]) => mg.c1 <= b && mg.c2 >= a));
  const covered = new Map<number, Range>();
  for (const mg of visMerges)
    for (let r = mg.r1; r <= Math.min(mg.r2, mg.r1 + 500); r++) for (let c = mg.c1; c <= Math.min(mg.c2, mg.c1 + 200); c++) covered.set(cellKey(r, c), mg);

  const drawRegion = (rr: [number, number], cc: [number, number]) => {
    const [r0, r1] = rr;
    const [c0, c1] = cc;
    if (r1 < r0 || c1 < c0) return;
    // pane-local coordinates, so a merged area crossing a frozen split continues into the next pane
    const frozenCols = c0 < vp.fc;
    const frozenRows = r0 < vp.fr;
    const X = (c: number) => (frozenCols ? vp.headerW + geo.cols.pos(c) * z : m.bodyX0 + (geo.cols.pos(c) - m.fcW - vp.scrollX) * z);
    const Y = (r: number) => (frozenRows ? vp.headerH + geo.rows.pos(r) * z : m.bodyY0 + (geo.rows.pos(r) - m.frH - vp.scrollY) * z);
    // the cell that paints a merged area in this pane: its anchor, or the pane's first cell of it when the anchor is outside
    const paintAt = (r: number, c: number): { r: number; c: number; mg?: Range } | null => {
      const mg = covered.get(cellKey(r, c));
      if (!mg) return { r, c };
      return r === Math.max(mg.r1, r0) && c === Math.max(mg.c1, c0) ? { r: mg.r1, c: mg.c1, mg } : null;
    };
    const clipX = Math.max(vp.headerW, X(c0));
    const clipY = Math.max(vp.headerH, Y(r0));
    const clipX2 = Math.min(W, X(c1 + 1));
    const clipY2 = Math.min(H, Y(r1 + 1));
    // frozen split: body regions must not paint over frozen panes
    const x0 = cc[0] >= vp.fc ? Math.max(clipX, m.bodyX0) : clipX;
    const y0 = rr[0] >= vp.fr ? Math.max(clipY, m.bodyY0) : clipY;
    const x1 = cc[0] < vp.fc ? Math.min(clipX2, m.bodyX0) : clipX2;
    const y1 = rr[0] < vp.fr ? Math.min(clipY2, m.bodyY0) : clipY2;
    if (x1 <= x0 || y1 <= y0) return;
    ctx.save();
    ctx.beginPath();
    ctx.rect(x0, y0, x1 - x0, y1 - y0);
    ctx.clip();

    // ---- gridlines
    if (sheet.view.showGrid) {
      ctx.strokeStyle = theme.line;
      ctx.lineWidth = 1 / dpr;
      ctx.beginPath();
      for (let c = c0; c <= c1 + 1; c++) {
        const x = snap(X(c) - 1 / dpr);
        ctx.moveTo(x, y0);
        ctx.lineTo(x, y1);
      }
      for (let r = r0; r <= r1 + 1; r++) {
        const y = snap(Y(r) - 1 / dpr);
        ctx.moveTo(x0, y);
        ctx.lineTo(x1, y);
      }
      ctx.stroke();
    }

    // ---- fills (cell fill, colour scales) & data bars
    for (let r = r0; r <= r1; r++) {
      const ry = Y(r);
      const rh = Y(r + 1) - ry;
      if (rh <= 0) continue;
      for (let c = c0; c <= c1; c++) {
        const at = paintAt(r, c);
        if (!at) continue;
        const mg = at.mg;
        const inf = info(at.r, at.c);
        const cx = X(at.c);
        const ty = mg ? Y(at.r) : ry;
        let cw = X(c + 1) - cx;
        let ch = rh;
        if (mg) {
          cw = X(mg.c2 + 1) - cx;
          ch = Y(mg.r2 + 1) - ty;
          // merged areas hide the gridlines inside them
          ctx.fillStyle = theme.bg;
          ctx.fillRect(cx, ty, cw - 1 / dpr, ch - 1 / dpr);
        }
        if (!inf || cw <= 0) continue;
        const fill = inf.cf?.scale ?? inf.st.fill;
        if (fill) {
          ctx.fillStyle = tone(fill, 'fill');
          ctx.fillRect(cx - 1 / dpr, ty - 1 / dpr, cw + 1 / dpr, ch + 1 / dpr);
        }
        if (inf.cf?.bar) {
          const b = inf.cf.bar;
          const pad = 2 * z;
          const bw = cw - pad * 2;
          const g = ctx.createLinearGradient(cx + pad + bw * b.from, 0, cx + pad + bw * b.to, 0);
          g.addColorStop(0, b.color);
          g.addColorStop(1, b.color + '40');
          ctx.fillStyle = b.negative ? b.color + 'b0' : g;
          ctx.fillRect(cx + pad + bw * b.from, ty + pad, Math.max(1, bw * (b.to - b.from)), ch - pad * 2);
        }
      }
    }

    // ---- text
    for (let r = r0; r <= r1; r++) {
      const ry = Y(r);
      const rh = Y(r + 1) - ry;
      if (rh <= 0) continue;
      // include cells left of the region whose text may overflow into view
      let start = c0;
      for (let k = 0; k < 30 && start > (c0 >= vp.fc ? vp.fc : 0); k++) {
        start--;
        if (info(r, start)?.text) break;
      }
      let end = c1;
      for (let k = 0; k < 30 && end < MAX_COLS - 1; k++) {
        end++;
        if (info(r, end)?.text) break;
      }
      for (let col = start; col <= end; col++) {
        const at = paintAt(r, col);
        if (!at) continue;
        const { r: ar, c, mg } = at;
        if (ov.editing && ov.editing.r === ar && ov.editing.c === c) continue;
        const inf = info(ar, c);
        if (!inf || !inf.text) continue;
        const st = inf.st;
        const cx = X(c);
        const ry = Y(ar);
        const cw = (mg ? X(mg.c2 + 1) : X(c + 1)) - cx;
        const ch = (mg ? Y(mg.r2 + 1) : Y(r + 1)) - ry;
        if (cw <= 0 || ch <= 0) continue;
        const { font, px } = cellFont(st, z);
        const pad = 3 * z + (st.indent ?? 0) * 9 * z;
        const isNum = typeof inf.value === 'number';
        const isBool = typeof inf.value === 'boolean' || isErr(inf.value);
        let align = st.hAlign ?? 'general';
        if (align === 'general') align = isNum && !showFormulas ? 'right' : isBool ? 'center' : 'left';
        if (align === 'centerContinuous' || align === 'distributed') align = 'center';
        let text = inf.text;
        const avail = cw - pad * 2;
        if (isNum && !st.wrap && !showFormulas) text = fitText(ctx, font, inf.value as number, st.numFmt, Math.max(4, cw - 4 * z), text);
        ctx.font = font;
        ctx.fillStyle = inf.color ? tone(namedColor(inf.color), 'text') : st.color ? tone(st.color, 'text') : theme.text;
        const metricsF = fontMetrics(ctx, font, px);
        const lineH = px * 1.22;
        const va = st.vAlign ?? 'bottom';
        const rot = st.rotation ?? 0;
        if (rot) {
          ctx.save();
          ctx.beginPath();
          ctx.rect(cx, ry, cw, ch);
          ctx.clip();
          ctx.translate(cx + cw / 2, ry + ch / 2);
          if (rot === 255) {
            let yy = -((text.length - 1) * lineH) / 2;
            for (const chr of text) {
              ctx.textAlign = 'center';
              ctx.textBaseline = 'middle';
              ctx.fillText(chr, 0, yy);
              yy += lineH;
            }
          } else {
            ctx.rotate((-rot * Math.PI) / 180);
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(text, 0, 0);
          }
          ctx.restore();
          continue;
        }
        if (st.wrap || align === 'justify') {
          const lines = wrapLines(ctx, font, text, Math.max(4, avail));
          const total = lines.length * lineH;
          let y = va === 'top' ? ry + 2 * z : va === 'middle' ? ry + (ch - total) / 2 : ry + ch - total - 2 * z;
          ctx.save();
          ctx.beginPath();
          ctx.rect(cx, ry, cw, ch);
          ctx.clip();
          ctx.textBaseline = 'alphabetic';
          for (const line of lines) {
            const lw = measure(ctx, font, line);
            const lx = align === 'right' ? cx + cw - pad - lw : align === 'center' ? cx + (cw - lw) / 2 : cx + pad;
            ctx.fillText(line, lx, y + (lineH + metricsF.ascent - metricsF.descent) / 2);
            if (st.underline || st.strike) decorate(ctx, st, lx, y + (lineH + metricsF.ascent - metricsF.descent) / 2, lw, px, dpr);
            y += lineH;
          }
          ctx.restore();
          continue;
        }
        const tw = measure(ctx, font, text);
        // horizontal position & overflow clip
        let clipL = cx;
        let clipR = cx + cw;
        let tx: number;
        if (align === 'fill') {
          const reps = Math.max(1, Math.floor(avail / Math.max(1, tw)));
          text = text.repeat(reps);
          tx = cx + pad;
        } else if (align === 'right') tx = cx + cw - pad - tw;
        else if (align === 'center') tx = cx + (cw - tw) / 2;
        else tx = cx + pad;
        const overflows = !mg && !isNum && !st.wrap && align !== 'fill' && tw > avail;
        if (overflows) {
          if (align === 'left' || align === 'center') {
            let cc = c + 1;
            while (cc < MAX_COLS && tx + tw > clipR && !info(r, cc)?.text && !covered.has(cellKey(r, cc))) {
              clipR = X(cc + 1);
              cc++;
              if (cc - c > 60) break;
            }
          }
          if (align === 'right' || align === 'center') {
            let cc = c - 1;
            while (cc >= 0 && tx < clipL && !info(r, cc)?.text && !covered.has(cellKey(r, cc))) {
              clipL = X(cc);
              cc--;
              if (c - cc > 60) break;
            }
          }
        }
        const baseY = va === 'top' ? ry + 2 * z + metricsF.ascent : va === 'middle' ? ry + ch / 2 + (metricsF.ascent - metricsF.descent) / 2 : ry + ch - 3 * z - metricsF.descent + 1 * z;
        ctx.save();
        ctx.beginPath();
        ctx.rect(clipL, ry, clipR - clipL, ch);
        ctx.clip();
        ctx.textBaseline = 'alphabetic';
        ctx.fillText(text, tx, baseY);
        if (st.underline || st.strike) decorate(ctx, st, tx, baseY, Math.min(tw, clipR - tx), px, dpr);
        ctx.restore();
      }
    }

    // ---- icons, notes
    for (let r = r0; r <= r1; r++) {
      const ry = Y(r);
      for (let c = c0; c <= c1; c++) {
        const inf = info(r, c);
        if (!inf) continue;
        const cx = X(c);
        const cw = X(c + 1) - cx;
        const ch = Y(r + 1) - ry;
        if (inf.cf?.icon) {
          const s = Math.min(ch - 4 * z, 14 * z);
          drawIcon(ctx, inf.cf.icon.set, inf.cf.icon.index, inf.cf.icon.total, cx + 3 * z, ry + (ch - s) / 2, s);
        }
        const cell = sheet.cells.get(cellKey(r, c));
        if (cell?.note) {
          ctx.fillStyle = '#e5484d';
          ctx.beginPath();
          ctx.moveTo(cx + cw - 7 * z, ry);
          ctx.lineTo(cx + cw, ry);
          ctx.lineTo(cx + cw, ry + 7 * z);
          ctx.fill();
        }
      }
    }

    // ---- borders (each edge drawn once, heavier style wins)
    ctx.lineCap = 'square';
    for (let r = r0; r <= r1 + 1; r++) {
      const y = Y(r) - 1 / dpr;
      for (let c = c0; c <= c1; c++) {
        const above = r > 0 ? info(r - 1, c)?.st.bb : undefined;
        const below = r < MAX_ROWS ? info(r, c)?.st.bt : undefined;
        const b = heavier(above, below);
        if (!b) continue;
        const inside = covered.get(cellKey(r, c));
        if (inside && inside.r1 < r && covered.get(cellKey(r - 1, c)) === inside) continue;
        strokeBorder(ctx, b, X(c) - 1 / dpr, y, X(c + 1) - 1 / dpr, y, dpr, b.color ? tone(b.color, 'line') : theme.text);
      }
    }
    for (let c = c0; c <= c1 + 1; c++) {
      const x = X(c) - 1 / dpr;
      for (let r = r0; r <= r1; r++) {
        const left = c > 0 ? info(r, c - 1)?.st.br : undefined;
        const right = info(r, c)?.st.bl;
        const b = heavier(left, right);
        if (!b) continue;
        const inside = covered.get(cellKey(r, c));
        if (inside && inside.c1 < c && covered.get(cellKey(r, c - 1)) === inside) continue;
        strokeBorder(ctx, b, x, Y(r) - 1 / dpr, x, Y(r + 1) - 1 / dpr, dpr, b.color ? tone(b.color, 'line') : theme.text);
      }
    }

    // ---- filter buttons
    const af = sheet.filter;
    if (af && af.range.r1 >= r0 && af.range.r1 <= r1) {
      for (let c = Math.max(af.range.c1, c0); c <= Math.min(af.range.c2, c1); c++) {
        const bx = X(c + 1) - 17 * z;
        const by = Y(af.range.r1) + (Y(af.range.r1 + 1) - Y(af.range.r1) - 15 * z) / 2;
        const active = !!af.columns[c - af.range.c1];
        drawFilterButton(ctx, bx, by, 15 * z, active, theme);
      }
    }

    ctx.restore();
  };

  for (const rr of rowsSets) for (const cc of colsSets) drawRegion(rr, cc);

  // ---- images/charts are DOM overlays; selection & overlays drawn on top
  ctx.save();
  ctx.beginPath();
  ctx.rect(vp.headerW, vp.headerH, W - vp.headerW, H - vp.headerH);
  ctx.clip();
  drawOverlay(ctx, m, sheet, ov, theme, z);
  ctx.restore();

  // ---- frozen pane lines
  if (vp.fr > 0) {
    ctx.fillStyle = theme.frozenLine;
    ctx.fillRect(0, m.bodyY0 - 1, W, 2 / dpr + 1);
  }
  if (vp.fc > 0) {
    ctx.fillStyle = theme.frozenLine;
    ctx.fillRect(m.bodyX0 - 1, 0, 2 / dpr + 1, H);
  }

  if (sheet.view.showHeaders) drawHeaders(ctx, m, sheet, ov, theme, dpr, rowsSets, colsSets, geo, vp);
  ctx.restore();
}

/* ============================================================ dark tones */

const toneCache = new Map<string, string>();

/**
 * Display-only colour mapping for dark mode: light fills turn dark and dark text turns light while
 * keeping hue and saturation, so sheets made in light mode stay readable. Files keep their colours.
 */
function darkTone(color: string, kind: 'fill' | 'text' | 'line'): string {
  const key = kind + color;
  const hit = toneCache.get(key);
  if (hit) return hit;
  const rgb = parseColor(color);
  let out = color;
  if (rgb) {
    const [h, s, l] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
    let l2 = l;
    if (kind === 'fill' && l > 0.5) l2 = 0.08 + (1 - l) * 0.85;
    else if (kind === 'text' && l < 0.5) l2 = 1 - l;
    else if (kind === 'line') l2 = Math.min(1, Math.max(0.3, 1 - l));
    if (l2 !== l) {
      const [r, g, b] = hslToRgb(h, s, l2);
      out = rgb[3] < 1 ? `rgba(${r},${g},${b},${rgb[3]})` : `rgb(${r},${g},${b})`;
    }
  }
  if (toneCache.size > 4000) toneCache.clear();
  toneCache.set(key, out);
  return out;
}

function parseColor(c: string): [number, number, number, number] | null {
  const hex = /^#([0-9a-f]{3,8})$/i.exec(c.trim());
  if (hex) {
    let h = hex[1];
    if (h.length === 3 || h.length === 4) h = [...h].map((x) => x + x).join('');
    if (h.length !== 6 && h.length !== 8) return null;
    const n = (i: number) => parseInt(h.slice(i, i + 2), 16);
    return [n(0), n(2), n(4), h.length === 8 ? n(6) / 255 : 1];
  }
  const fn = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.]+%?))?\s*\)$/i.exec(c.trim());
  if (fn) {
    const a = fn[4] === undefined ? 1 : fn[4].endsWith('%') ? parseFloat(fn[4]) / 100 : parseFloat(fn[4]);
    return [+fn[1], +fn[2], +fn[3], a];
  }
  return null;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
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

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) {
    const v = Math.round(l * 255);
    return [v, v, v];
  }
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
  return [Math.round(f(h + 1 / 3) * 255), Math.round(f(h) * 255), Math.round(f(h - 1 / 3) * 255)];
}

function namedColor(c: string): string {
  const map: Record<string, string> = { red: '#e00000', blue: '#0000ff', green: '#00a000', yellow: '#c8a000', magenta: '#d000d0', cyan: '#00a0a0', white: '#ffffff', black: '#000000' };
  return map[c.toLowerCase()] ?? c;
}

function decorate(ctx: CanvasRenderingContext2D, st: CellStyle, x: number, baseY: number, w: number, px: number, dpr: number) {
  ctx.save();
  ctx.strokeStyle = ctx.fillStyle as string;
  ctx.lineWidth = Math.max(1 / dpr, px / 16);
  ctx.beginPath();
  if (st.underline) {
    const y = baseY + px * 0.12;
    ctx.moveTo(x, y);
    ctx.lineTo(x + w, y);
    if (st.underline === 'double') {
      ctx.moveTo(x, y + px * 0.12);
      ctx.lineTo(x + w, y + px * 0.12);
    }
  }
  if (st.strike) {
    const y = baseY - px * 0.3;
    ctx.moveTo(x, y);
    ctx.lineTo(x + w, y);
  }
  ctx.stroke();
  ctx.restore();
}

function drawFilterButton(ctx: CanvasRenderingContext2D, x: number, y: number, s: number, active: boolean, theme: GridTheme) {
  ctx.save();
  ctx.fillStyle = active ? theme.accent : theme.headerBg;
  ctx.strokeStyle = active ? theme.accent : theme.headerLine;
  ctx.lineWidth = 1;
  roundRect(ctx, x, y, s, s, 3);
  ctx.fill();
  ctx.stroke();
  ctx.fillStyle = active ? '#fff' : theme.headerFg;
  ctx.beginPath();
  if (active) {
    // funnel
    ctx.moveTo(x + s * 0.2, y + s * 0.28);
    ctx.lineTo(x + s * 0.8, y + s * 0.28);
    ctx.lineTo(x + s * 0.56, y + s * 0.55);
    ctx.lineTo(x + s * 0.56, y + s * 0.78);
    ctx.lineTo(x + s * 0.44, y + s * 0.7);
    ctx.lineTo(x + s * 0.44, y + s * 0.55);
  } else {
    ctx.moveTo(x + s * 0.28, y + s * 0.4);
    ctx.lineTo(x + s * 0.72, y + s * 0.4);
    ctx.lineTo(x + s * 0.5, y + s * 0.65);
  }
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

function roundRect(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

/* ============================================================= overlay */

function mergedRect(sheet: Sheet, r: number, c: number): Range {
  return sheet.mergeAt(r, c) ?? { r1: r, c1: c, r2: r, c2: c };
}

function expandForMerges(sheet: Sheet, rg: Range): Range {
  let out = { ...rg };
  for (let guard = 0; guard < 8; guard++) {
    let changed = false;
    for (const mg of sheet.merges) {
      if (mg.r1 <= out.r2 && mg.r2 >= out.r1 && mg.c1 <= out.c2 && mg.c2 >= out.c1) {
        const n = { r1: Math.min(out.r1, mg.r1), c1: Math.min(out.c1, mg.c1), r2: Math.max(out.r2, mg.r2), c2: Math.max(out.c2, mg.c2) };
        if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) {
          out = n;
          changed = true;
        }
      }
    }
    if (!changed) break;
  }
  return out;
}

function drawOverlay(ctx: CanvasRenderingContext2D, m: Mapper, sheet: Sheet, ov: Overlay, theme: GridTheme, z: number) {
  const sel = ov.sel;
  const clampRect = (rg: Range) => {
    const r = m.rect(rg);
    // keep huge selections within drawable limits
    const x = Math.max(r.x, -10);
    const y = Math.max(r.y, -10);
    return { x, y, w: Math.min(r.x + r.w, m.vp.width + 10) - x, h: Math.min(r.y + r.h, m.vp.height + 10) - y };
  };

  // spill range of the active cell
  if (ov.spillHint) {
    const r = clampRect(ov.spillHint);
    ctx.save();
    ctx.strokeStyle = theme.accent;
    ctx.setLineDash([4, 3]);
    ctx.lineWidth = 1;
    ctx.strokeRect(Math.round(r.x) + 0.5, Math.round(r.y) + 0.5, r.w - 2, r.h - 2);
    ctx.restore();
  }

  // formula reference highlights
  for (const ref of ov.refs ?? []) {
    const r = clampRect(expandForMerges(sheet, ref.range));
    ctx.fillStyle = ref.color + '1f';
    ctx.fillRect(r.x, r.y, r.w, r.h);
    ctx.strokeStyle = ref.color;
    ctx.lineWidth = 2;
    ctx.strokeRect(r.x + 1, r.y + 1, r.w - 3, r.h - 3);
    ctx.fillStyle = ref.color;
    const hs = 5;
    for (const [hx, hy] of [
      [r.x, r.y],
      [r.x + r.w - 2, r.y],
      [r.x, r.y + r.h - 2],
      [r.x + r.w - 2, r.y + r.h - 2],
    ])
      ctx.fillRect(hx - hs / 2, hy - hs / 2, hs, hs);
  }

  if (ov.findHits?.size) {
    ctx.fillStyle = '#ffd60a55';
    for (const k of ov.findHits) {
      const r = Math.floor(k / MAX_COLS);
      const c = k % MAX_COLS;
      const rect = m.rect({ r1: r, c1: c, r2: r, c2: c });
      if (rect.x > m.vp.width || rect.y > m.vp.height || rect.x + rect.w < 0 || rect.y + rect.h < 0) continue;
      ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
    }
  }

  // selection fill (the active cell keeps its normal background, like Excel)
  const accent = theme.accent;
  const single = sel.ranges.length === 1;
  const act = mergedRect(sheet, sel.active.r, sel.active.c);
  const ar = m.rect(act);
  ctx.fillStyle = theme.selFill;
  for (const rg0 of sel.ranges) {
    const rg = expandForMerges(sheet, rg0);
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && single) continue;
    const r = clampRect(rg);
    const holeX1 = Math.max(r.x, ar.x);
    const holeY1 = Math.max(r.y, ar.y);
    const holeX2 = Math.min(r.x + r.w, ar.x + ar.w);
    const holeY2 = Math.min(r.y + r.h, ar.y + ar.h);
    if (holeX2 <= holeX1 || holeY2 <= holeY1) {
      ctx.fillRect(r.x, r.y, r.w, r.h);
      continue;
    }
    ctx.fillRect(r.x, r.y, r.w, holeY1 - r.y);
    ctx.fillRect(r.x, holeY2, r.w, r.y + r.h - holeY2);
    ctx.fillRect(r.x, holeY1, holeX1 - r.x, holeY2 - holeY1);
    ctx.fillRect(holeX2, holeY1, r.x + r.w - holeX2, holeY2 - holeY1);
  }

  // selection borders
  const lw = Math.max(2, Math.round(2 * Math.min(1.5, z)));
  ctx.strokeStyle = ov.focused ? accent : accent + 'aa';
  for (let i = 0; i < sel.ranges.length; i++) {
    const rg = expandForMerges(sheet, sel.ranges[i]);
    const r = clampRect(rg);
    const primary = i === sel.ranges.length - 1;
    ctx.lineWidth = primary && single ? lw : 1;
    const inset = primary && single ? lw / 2 - 0.5 : 0.5;
    ctx.strokeRect(Math.round(r.x) + inset - 1, Math.round(r.y) + inset - 1, Math.round(r.w) - inset * 2 + 1, Math.round(r.h) - inset * 2 + 1);
    if (primary && single && !ov.editing) {
      // fill handle
      const hs = Math.max(6, Math.round(6 * Math.min(1.4, z)));
      const hx = Math.round(r.x + r.w) - hs / 2 - 1;
      const hy = Math.round(r.y + r.h) - hs / 2 - 1;
      ctx.fillStyle = theme.bg;
      ctx.fillRect(hx - 1, hy - 1, hs + 2, hs + 2);
      ctx.fillStyle = accent;
      ctx.fillRect(hx, hy, hs, hs);
    }
  }
  if (!single) {
    ctx.lineWidth = 2;
    ctx.strokeStyle = accent;
    ctx.strokeRect(ar.x + 0.5, ar.y + 0.5, ar.w - 2, ar.h - 2);
  }

  // copy marquee
  if (ov.clip) {
    const r = clampRect(ov.clip.range);
    ctx.save();
    ctx.lineWidth = 2;
    ctx.strokeStyle = accent;
    ctx.setLineDash([5, 4]);
    ctx.lineDashOffset = -ov.clip.phase;
    ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 2, r.h - 2);
    ctx.restore();
  }

  for (const pv of [ov.fillPreview, ov.movePreview]) {
    if (!pv) continue;
    const r = clampRect(pv);
    ctx.save();
    ctx.lineWidth = 1.5;
    ctx.strokeStyle = theme.dark ? '#c9cbd4' : '#444';
    ctx.setLineDash([4, 3]);
    ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 2, r.h - 2);
    ctx.restore();
  }

  // validation dropdown arrow on the active cell
  const v = sheet.validations.find((x) => x.type === 'list' && x.showDropdown !== false && x.ranges.some((rg) => sel.active.r >= rg.r1 && sel.active.r <= rg.r2 && sel.active.c >= rg.c1 && sel.active.c <= rg.c2));
  if (v && !ov.editing) {
    const s = Math.min(18 * z, ar.h - 2);
    const x = ar.x + ar.w + 2;
    const y = ar.y + ar.h - s - 1;
    ctx.fillStyle = theme.headerBg;
    ctx.strokeStyle = theme.headerLine;
    ctx.fillRect(x, y, s, s);
    ctx.strokeRect(x + 0.5, y + 0.5, s - 1, s - 1);
    ctx.fillStyle = theme.headerFg;
    ctx.beginPath();
    ctx.moveTo(x + s * 0.3, y + s * 0.42);
    ctx.lineTo(x + s * 0.7, y + s * 0.42);
    ctx.lineTo(x + s * 0.5, y + s * 0.64);
    ctx.fill();
  }
}

/* ============================================================= headers */

function drawHeaders(
  ctx: CanvasRenderingContext2D,
  m: Mapper,
  sheet: Sheet,
  ov: Overlay,
  theme: GridTheme,
  dpr: number,
  rowsSets: [number, number][],
  colsSets: [number, number][],
  geo: SheetGeometry,
  vp: Viewport,
) {
  const z = vp.zoom;
  const W = vp.width;
  const H = vp.height;
  const sel = ov.sel;
  const selCols = (c: number) => sel.ranges.some((rg) => c >= rg.c1 && c <= rg.c2);
  const selRows = (r: number) => sel.ranges.some((rg) => r >= rg.r1 && r <= rg.r2);
  const fullCol = (c: number) => sel.ranges.some((rg) => rg.r1 === 0 && rg.r2 === MAX_ROWS - 1 && c >= rg.c1 && c <= rg.c2);
  const fullRow = (r: number) => sel.ranges.some((rg) => rg.c1 === 0 && rg.c2 === MAX_COLS - 1 && r >= rg.r1 && r <= rg.r2);
  const font = `${Math.max(9, 11 * Math.min(1.3, Math.max(0.8, z)))}px ${theme.uiFont}`;

  // column header
  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(0, 0, W, vp.headerH);
  ctx.fillRect(0, 0, vp.headerW, H);
  ctx.font = font;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  for (const [c0, c1] of colsSets) {
    ctx.save();
    ctx.beginPath();
    const x0 = c0 >= vp.fc ? m.bodyX0 : vp.headerW;
    ctx.rect(x0, 0, W - x0, vp.headerH);
    ctx.clip();
    for (let c = c0; c <= c1; c++) {
      const x = m.x(c);
      const w = m.x(c + 1) - x;
      if (w <= 0) continue;
      const selected = selCols(c);
      if (selected) {
        ctx.fillStyle = fullCol(c) ? theme.accent : theme.headerSelBg;
        ctx.fillRect(x, 0, w, vp.headerH);
        if (!fullCol(c)) {
          ctx.fillStyle = theme.accent;
          ctx.fillRect(x, vp.headerH - 2, w, 2);
        }
      }
      ctx.fillStyle = fullCol(c) ? '#fff' : selected ? theme.headerSelFg : theme.headerFg;
      if (w > 8) ctx.fillText(colName(c), x + w / 2, vp.headerH / 2 + 0.5);
      ctx.fillStyle = theme.headerLine;
      ctx.fillRect(Math.round(x + w) - 1, 0, 1 / dpr, vp.headerH);
      // hidden column marker
      if (c + 1 < MAX_COLS && geo.cols.sizeOf(c + 1) === 0) {
        ctx.fillStyle = theme.accent;
        ctx.fillRect(Math.round(x + w) - 2, 3, 2, vp.headerH - 6);
      }
    }
    ctx.restore();
  }
  // row header
  for (const [r0, r1] of rowsSets) {
    ctx.save();
    ctx.beginPath();
    const y0 = r0 >= vp.fr ? m.bodyY0 : vp.headerH;
    ctx.rect(0, y0, vp.headerW, H - y0);
    ctx.clip();
    for (let r = r0; r <= r1; r++) {
      const y = m.y(r);
      const h = m.y(r + 1) - y;
      if (h <= 0) continue;
      const selected = selRows(r);
      if (selected) {
        ctx.fillStyle = fullRow(r) ? theme.accent : theme.headerSelBg;
        ctx.fillRect(0, y, vp.headerW, h);
        if (!fullRow(r)) {
          ctx.fillStyle = theme.accent;
          ctx.fillRect(vp.headerW - 2, y, 2, h);
        }
      }
      ctx.fillStyle = fullRow(r) ? '#fff' : selected ? theme.headerSelFg : sheet.filterHidden.has(r + 1) ? theme.accent : theme.headerFg;
      if (h > 8) ctx.fillText(String(r + 1), vp.headerW / 2, y + h / 2 + 0.5);
      ctx.fillStyle = theme.headerLine;
      ctx.fillRect(0, Math.round(y + h) - 1, vp.headerW, 1 / dpr);
      if (r + 1 < MAX_ROWS && geo.rows.sizeOf(r + 1) === 0) {
        ctx.fillStyle = theme.accent;
        ctx.fillRect(3, Math.round(y + h) - 2, vp.headerW - 6, 2);
      }
    }
    ctx.restore();
  }
  // borders & corner
  ctx.fillStyle = theme.headerLine;
  ctx.fillRect(0, vp.headerH - 1, W, 1);
  ctx.fillRect(vp.headerW - 1, 0, 1, H);
  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(0, 0, vp.headerW - 1, vp.headerH - 1);
  ctx.fillStyle = theme.headerLine;
  ctx.beginPath();
  ctx.moveTo(vp.headerW - 4, vp.headerH - 12);
  ctx.lineTo(vp.headerW - 4, vp.headerH - 4);
  ctx.lineTo(vp.headerW - 12, vp.headerH - 4);
  ctx.fill();
}

export function headerWidth(maxRow: number, zoom: number): number {
  const digits = String(Math.max(100, maxRow + 1)).length;
  return Math.round((digits * 7.5 + 16) * Math.min(1.3, Math.max(0.8, zoom)));
}
