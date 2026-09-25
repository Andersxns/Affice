/**
 * Preset shape geometry (a practical subset of PowerPoint's prstGeom set) rendered as SVG paths.
 * Adjust values use PowerPoint's units (1/100000 of the shorter side or of the width/height), so shapes
 * imported from .pptx keep their proportions and export back unchanged. Unknown presets draw as a
 * rectangle but keep their name, so PowerPoint still shows the original shape.
 */
import type { CustomPath } from './model';

export type PathFill = 'norm' | 'none' | 'darken' | 'darkenLess' | 'lighten' | 'lightenLess';

export interface GeomPath {
  d: string;
  fill: PathFill;
  stroke: boolean;
}

export interface Geometry {
  paths: GeomPath[];
  /** Text rectangle as l, t, r, b in px from the shape's edges. */
  text?: [number, number, number, number];
}

class PathBuilder {
  s = '';
  x = 0;
  y = 0;
  m(x: number, y: number) {
    this.s += `M${r(x)} ${r(y)}`;
    this.x = x;
    this.y = y;
    return this;
  }
  l(x: number, y: number) {
    this.s += `L${r(x)} ${r(y)}`;
    this.x = x;
    this.y = y;
    return this;
  }
  c(x1: number, y1: number, x2: number, y2: number, x: number, y: number) {
    this.s += `C${r(x1)} ${r(y1)} ${r(x2)} ${r(y2)} ${r(x)} ${r(y)}`;
    this.x = x;
    this.y = y;
    return this;
  }
  q(x1: number, y1: number, x: number, y: number) {
    this.s += `Q${r(x1)} ${r(y1)} ${r(x)} ${r(y)}`;
    this.x = x;
    this.y = y;
    return this;
  }
  /** DrawingML arcTo: ellipse radii, start angle and sweep (degrees, clockwise), from the current point. */
  arc(wR: number, hR: number, stAng: number, swAng: number) {
    if (wR <= 0 || hR <= 0 || !swAng) return this;
    const st = (stAng * Math.PI) / 180;
    // centre of the ellipse such that the current point lies at stAng on it
    const cx = this.x - wR * Math.cos(st);
    const cy = this.y - hR * Math.sin(st);
    const draw = (from: number, sweep: number) => {
      const end = ((from + sweep) * Math.PI) / 180;
      const ex = cx + wR * Math.cos(end);
      const ey = cy + hR * Math.sin(end);
      const large = Math.abs(sweep) > 180 ? 1 : 0;
      const sweepFlag = sweep > 0 ? 1 : 0;
      this.s += `A${r(wR)} ${r(hR)} 0 ${large} ${sweepFlag} ${r(ex)} ${r(ey)}`;
      this.x = ex;
      this.y = ey;
    };
    if (Math.abs(swAng) >= 360) {
      draw(stAng, swAng / 2);
      draw(stAng + swAng / 2, swAng / 2);
    } else draw(stAng, swAng);
    return this;
  }
  z() {
    this.s += 'Z';
    return this;
  }
  toString() {
    return this.s;
  }
}

const r = (v: number) => Math.round(v * 100) / 100;
const P = () => new PathBuilder();

function poly(pts: [number, number][]): string {
  const p = P().m(pts[0][0], pts[0][1]);
  for (let i = 1; i < pts.length; i++) p.l(pts[i][0], pts[i][1]);
  return p.z().toString();
}

function ellipsePath(x: number, y: number, w: number, h: number): string {
  const rx = w / 2;
  const ry = h / 2;
  return `M${r(x)} ${r(y + ry)}A${r(rx)} ${r(ry)} 0 1 1 ${r(x + w)} ${r(y + ry)}A${r(rx)} ${r(ry)} 0 1 1 ${r(x)} ${r(y + ry)}Z`;
}

function roundRectPath(x: number, y: number, w: number, h: number, rad: number): string {
  const rr = Math.max(0, Math.min(rad, w / 2, h / 2));
  if (!rr) return poly([[x, y], [x + w, y], [x + w, y + h], [x, y + h]]);
  return P()
    .m(x + rr, y)
    .l(x + w - rr, y)
    .arc(rr, rr, 270, 90)
    .l(x + w, y + h - rr)
    .arc(rr, rr, 0, 90)
    .l(x + rr, y + h)
    .arc(rr, rr, 90, 90)
    .l(x, y + rr)
    .arc(rr, rr, 180, 90)
    .z()
    .toString();
}

function regularPolygon(n: number, w: number, h: number, rotDeg = -90): [number, number][] {
  const pts: [number, number][] = [];
  for (let i = 0; i < n; i++) {
    const a = ((rotDeg + (360 / n) * i) * Math.PI) / 180;
    pts.push([w / 2 + (w / 2) * Math.cos(a), h / 2 + (h / 2) * Math.sin(a)]);
  }
  return pts;
}

function starPoints(n: number, w: number, h: number, inner: number): [number, number][] {
  const pts: [number, number][] = [];
  for (let i = 0; i < n * 2; i++) {
    const a = ((-90 + (180 / n) * i) * Math.PI) / 180;
    const k = i % 2 ? inner : 1;
    pts.push([w / 2 + (w / 2) * k * Math.cos(a), h / 2 + (h / 2) * k * Math.sin(a)]);
  }
  return pts;
}

const one = (d: string, text?: Geometry['text']): Geometry => ({ paths: [{ d, fill: 'norm', stroke: true }], text });

/** Default adjust values (PowerPoint defaults). */
export const DEFAULT_ADJ: Record<string, Record<string, number>> = {
  roundRect: { adj: 16667 },
  triangle: { adj: 50000 },
  parallelogram: { adj: 25000 },
  trapezoid: { adj: 25000 },
  hexagon: { adj: 25000 },
  octagon: { adj: 29289 },
  plus: { adj: 25000 },
  donut: { adj: 25000 },
  frame: { adj1: 12500 },
  can: { adj: 25000 },
  cube: { adj: 25000 },
  chevron: { adj: 50000 },
  homePlate: { adj: 50000 },
  rightArrow: { adj1: 50000, adj2: 50000 },
  leftArrow: { adj1: 50000, adj2: 50000 },
  upArrow: { adj1: 50000, adj2: 50000 },
  downArrow: { adj1: 50000, adj2: 50000 },
  leftRightArrow: { adj1: 50000, adj2: 50000 },
  upDownArrow: { adj1: 50000, adj2: 50000 },
  notchedRightArrow: { adj1: 50000, adj2: 50000 },
  wedgeRectCallout: { adj1: -20833, adj2: 62500 },
  wedgeRoundRectCallout: { adj1: -20833, adj2: 62500, adj3: 16667 },
  wedgeEllipseCallout: { adj1: -20833, adj2: 62500 },
  star4: { adj: 12500 },
  star5: { adj: 19098 },
  star6: { adj: 28868 },
  star7: { adj: 34601 },
  star8: { adj: 37500 },
  star10: { adj: 42533 },
  star12: { adj: 37500 },
  star16: { adj: 37500 },
  star24: { adj: 37500 },
  snip1Rect: { adj: 16667 },
  snip2SameRect: { adj1: 16667, adj2: 0 },
  round1Rect: { adj: 16667 },
  round2SameRect: { adj1: 16667, adj2: 0 },
  bentConnector3: { adj1: 50000 },
  curvedConnector3: { adj1: 50000 },
  foldedCorner: { adj: 16667 },
  teardrop: { adj: 100000 },
  pie: { adj1: 0, adj2: 16200000 },
  blockArc: { adj1: 10800000, adj2: 0, adj3: 25000 },
  moon: { adj: 50000 },
  smileyFace: { adj: 4653 },
  wave: { adj1: 12500, adj2: 0 },
  doubleWave: { adj1: 6250, adj2: 0 },
  bracketPair: { adj: 16667 },
  bracePair: { adj: 8333 },
  leftBracket: { adj: 8333 },
  rightBracket: { adj: 8333 },
  leftBrace: { adj1: 8333, adj2: 50000 },
  rightBrace: { adj1: 8333, adj2: 50000 },
  plaque: { adj: 16667 },
  bevel: { adj: 12500 },
  halfFrame: { adj1: 33333, adj2: 33333 },
  corner: { adj1: 50000, adj2: 50000 },
  diagStripe: { adj: 50000 },
  mathPlus: { adj1: 23520 },
  mathMinus: { adj1: 23520 },
  mathMultiply: { adj1: 23520 },
  mathEqual: { adj1: 23520, adj2: 11760 },
};

const LINES = new Set(['line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'curvedConnector2', 'curvedConnector3', 'curvedConnector4', 'arc', 'bracketPair', 'bracePair', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace']);

/** Shapes that are drawn as strokes only (no fill by default). */
export function isLineGeom(geom: string): boolean {
  return LINES.has(geom);
}

export function isConnector(geom: string): boolean {
  return geom === 'line' || geom.startsWith('straightConnector') || geom.startsWith('bentConnector') || geom.startsWith('curvedConnector');
}

export function presetGeometry(geom: string, w: number, h: number, adjIn?: Record<string, number>): Geometry {
  const adj = { ...(DEFAULT_ADJ[geom] ?? {}), ...(adjIn ?? {}) };
  const ss = Math.min(w, h);
  const a = (k = 'adj') => (adj[k] ?? 0) / 100000;
  switch (geom) {
    case 'rect':
    case 'flowChartProcess':
    case 'textBox':
      return one(poly([[0, 0], [w, 0], [w, h], [0, h]]));
    case 'roundRect':
    case 'flowChartAlternateProcess': {
      const rad = ss * (geom === 'roundRect' ? Math.min(0.5, a()) : 0.1667);
      const ins = rad * 0.29289;
      return one(roundRectPath(0, 0, w, h, rad), [ins, ins, ins, ins]);
    }
    case 'ellipse':
    case 'flowChartConnector': {
      const ix = w * 0.14645;
      const iy = h * 0.14645;
      return one(ellipsePath(0, 0, w, h), [ix, iy, ix, iy]);
    }
    case 'triangle':
    case 'flowChartExtract': {
      const x = w * (geom === 'triangle' ? a() : 0.5);
      return one(poly([[x, 0], [w, h], [0, h]]), [x / 2, h / 2, (w - x) / 2, 0]);
    }
    case 'flowChartMerge':
      return one(poly([[0, 0], [w, 0], [w / 2, h]]), [w / 4, 0, w / 4, h / 2]);
    case 'rtTriangle':
      return one(poly([[0, 0], [w, h], [0, h]]), [w / 12, (h * 7) / 12, (w * 7) / 12, h / 12]);
    case 'diamond':
    case 'flowChartDecision':
      return one(poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]), [w / 4, h / 4, w / 4, h / 4]);
    case 'parallelogram':
    case 'flowChartInputOutput': {
      const x = geom === 'parallelogram' ? ss * a() : w / 5;
      return one(poly([[x, 0], [w, 0], [w - x, h], [0, h]]), [x / 2 + w / 12, h / 12, x / 2 + w / 12, h / 12]);
    }
    case 'trapezoid':
    case 'flowChartManualOperation': {
      const x = geom === 'trapezoid' ? ss * a() : w / 5;
      return geom === 'trapezoid' ? one(poly([[x, 0], [w - x, 0], [w, h], [0, h]]), [x / 2, h / 6, x / 2, 0]) : one(poly([[0, 0], [w, 0], [w - x, h], [x, h]]), [x, 0, x, 0]);
    }
    case 'pentagon':
      return one(poly(regularPolygon(5, w, h)), [w * 0.19, h * 0.3, w * 0.19, h * 0.05]);
    case 'hexagon':
    case 'flowChartPreparation': {
      const x = geom === 'hexagon' ? ss * a() : w / 5;
      return one(poly([[0, h / 2], [x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h]]), [x * 0.6, h * 0.1, x * 0.6, h * 0.1]);
    }
    case 'heptagon':
      return one(poly(regularPolygon(7, w, h)), [w * 0.15, h * 0.15, w * 0.15, h * 0.1]);
    case 'octagon': {
      const x = ss * a();
      return one(poly([[x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]]), [x / 2, x / 2, x / 2, x / 2]);
    }
    case 'decagon':
      return one(poly(regularPolygon(10, w, h, -90 + 18)), [w * 0.1, h * 0.1, w * 0.1, h * 0.1]);
    case 'dodecagon':
      return one(poly(regularPolygon(12, w, h, -90 + 15)), [w * 0.1, h * 0.1, w * 0.1, h * 0.1]);
    case 'star4':
    case 'star5':
    case 'star6':
    case 'star7':
    case 'star8':
    case 'star10':
    case 'star12':
    case 'star16':
    case 'star24': {
      const n = parseInt(geom.slice(4), 10);
      const inner = Math.max(0.05, Math.min(0.95, a() * 2));
      return one(poly(starPoints(n, w, h, inner)), [w * 0.3, h * 0.3, w * 0.3, h * 0.3]);
    }
    case 'rightArrow':
    case 'leftArrow': {
      const th = h * a('adj1');
      const hl = Math.min(w, ss * a('adj2'));
      const y1 = (h - th) / 2;
      const pts: [number, number][] = [[0, y1], [w - hl, y1], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h - y1], [0, h - y1]];
      return one(poly(geom === 'rightArrow' ? pts : pts.map(([x, y]) => [w - x, y])), geom === 'rightArrow' ? [0, y1, hl / 2, y1] : [hl / 2, y1, 0, y1]);
    }
    case 'upArrow':
    case 'downArrow': {
      const th = w * a('adj1');
      const hl = Math.min(h, ss * a('adj2'));
      const x1 = (w - th) / 2;
      const pts: [number, number][] = [[x1, h], [x1, hl], [0, hl], [w / 2, 0], [w, hl], [w - x1, hl], [w - x1, h]];
      return one(poly(geom === 'upArrow' ? pts : pts.map(([x, y]) => [x, h - y])), geom === 'upArrow' ? [x1, hl / 2, x1, 0] : [x1, 0, x1, hl / 2]);
    }
    case 'leftRightArrow': {
      const th = h * a('adj1');
      const hl = Math.min(w / 2, ss * a('adj2'));
      const y1 = (h - th) / 2;
      return one(poly([[0, h / 2], [hl, 0], [hl, y1], [w - hl, y1], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h - y1], [hl, h - y1], [hl, h]]), [hl / 2, y1, hl / 2, y1]);
    }
    case 'upDownArrow': {
      const th = w * a('adj1');
      const hl = Math.min(h / 2, ss * a('adj2'));
      const x1 = (w - th) / 2;
      return one(poly([[w / 2, 0], [w, hl], [w - x1, hl], [w - x1, h - hl], [w, h - hl], [w / 2, h], [0, h - hl], [x1, h - hl], [x1, hl], [0, hl]]), [x1, hl / 2, x1, hl / 2]);
    }
    case 'notchedRightArrow': {
      const th = h * a('adj1');
      const hl = Math.min(w, ss * a('adj2'));
      const y1 = (h - th) / 2;
      const notch = (th / 2) * (hl / (h / 2));
      return one(poly([[0, y1], [w - hl, y1], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h - y1], [0, h - y1], [Math.min(notch, w - hl), h / 2]]), [notch, y1, hl / 2, y1]);
    }
    case 'chevron': {
      const x = Math.min(w, ss * a());
      return one(poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]]), [x, 0, x, 0]);
    }
    case 'homePlate': {
      const x = Math.min(w, ss * a());
      return one(poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h]]), [0, 0, x / 2, 0]);
    }
    case 'plus':
    case 'flowChartOr': {
      const x = geom === 'plus' ? ss * a() : 0;
      if (geom === 'flowChartOr') return { paths: [{ d: ellipsePath(0, 0, w, h), fill: 'norm', stroke: true }, { d: `M${w / 2} 0L${w / 2} ${h}M0 ${h / 2}L${w} ${h / 2}`, fill: 'none', stroke: true }] };
      return one(poly([[x, 0], [w - x, 0], [w - x, x], [w, x], [w, h - x], [w - x, h - x], [w - x, h], [x, h], [x, h - x], [0, h - x], [0, x], [x, x]]), [0, x, 0, x]);
    }
    case 'donut': {
      const t = ss * a();
      return { paths: [{ d: ellipsePath(0, 0, w, h) + ellipsePath(t, t, w - 2 * t, h - 2 * t), fill: 'norm', stroke: true }], text: [w * 0.15, h * 0.15, w * 0.15, h * 0.15] };
    }
    case 'frame': {
      const t = ss * a('adj1');
      return { paths: [{ d: poly([[0, 0], [w, 0], [w, h], [0, h]]) + poly([[t, t], [t, h - t], [w - t, h - t], [w - t, t]]), fill: 'norm', stroke: true }], text: [t, t, t, t] };
    }
    case 'can':
    case 'flowChartMagneticDisk': {
      const eh = geom === 'can' ? ss * a() : h / 6;
      const ry = eh / 2;
      const body = P().m(0, ry).arc(w / 2, ry, 180, -180).l(w, h - ry).arc(w / 2, ry, 0, 180).z().toString();
      const top = ellipsePath(0, 0, w, eh);
      return { paths: [{ d: body, fill: 'norm', stroke: true }, { d: top, fill: 'lighten', stroke: true }], text: [0, eh, 0, ry] };
    }
    case 'cube': {
      const d = ss * a();
      return {
        paths: [
          { d: poly([[0, d], [w - d, d], [w - d, h], [0, h]]), fill: 'norm', stroke: true },
          { d: poly([[0, d], [d, 0], [w, 0], [w - d, d]]), fill: 'lightenLess', stroke: true },
          { d: poly([[w - d, d], [w, 0], [w, h - d], [w - d, h]]), fill: 'darkenLess', stroke: true },
        ],
        text: [0, d, d, 0],
      };
    }
    case 'bevel': {
      const d = ss * a();
      return {
        paths: [
          { d: poly([[d, d], [w - d, d], [w - d, h - d], [d, h - d]]), fill: 'norm', stroke: true },
          { d: poly([[0, 0], [w, 0], [w - d, d], [d, d]]), fill: 'lightenLess', stroke: true },
          { d: poly([[w, 0], [w, h], [w - d, h - d], [w - d, d]]), fill: 'darkenLess', stroke: true },
          { d: poly([[0, h], [d, h - d], [w - d, h - d], [w, h]]), fill: 'darken', stroke: true },
          { d: poly([[0, 0], [d, d], [d, h - d], [0, h]]), fill: 'lighten', stroke: true },
        ],
        text: [d, d, d, d],
      };
    }
    case 'wedgeRectCallout':
    case 'wedgeRoundRectCallout':
    case 'wedgeEllipseCallout': {
      const tx = w / 2 + w * a('adj1');
      const ty = h / 2 + h * a('adj2');
      if (geom === 'wedgeEllipseCallout') {
        // ellipse with a wedge towards the tip
        const ang = Math.atan2((ty - h / 2) / (h / 2), (tx - w / 2) / (w / 2));
        const spread = 0.2;
        const p1 = [w / 2 + (w / 2) * Math.cos(ang - spread), h / 2 + (h / 2) * Math.sin(ang - spread)];
        const p2 = [w / 2 + (w / 2) * Math.cos(ang + spread), h / 2 + (h / 2) * Math.sin(ang + spread)];
        const d = `M${r(p2[0])} ${r(p2[1])}A${r(w / 2)} ${r(h / 2)} 0 1 1 ${r(p1[0])} ${r(p1[1])}L${r(tx)} ${r(ty)}Z`;
        return one(d, [w * 0.15, h * 0.15, w * 0.15, h * 0.15]);
      }
      const rad = geom === 'wedgeRoundRectCallout' ? ss * a('adj3') : 0;
      // wedge base on the side facing the tip
      const dx = (tx - w / 2) / w;
      const dy = (ty - h / 2) / h;
      const p = P();
      const base = (s: number, e: number) => [s, e];
      if (Math.abs(dy) >= Math.abs(dx) && dy > 0) {
        const [s, e] = base(w * (tx < w / 2 ? 0.2 : 0.5), w * (tx < w / 2 ? 0.35 : 0.65));
        p.m(rad, 0).l(w - rad, 0);
        if (rad) p.arc(rad, rad, 270, 90);
        p.l(w, h - rad);
        if (rad) p.arc(rad, rad, 0, 90);
        p.l(e, h).l(tx, ty).l(s, h).l(rad, h);
        if (rad) p.arc(rad, rad, 90, 90);
        p.l(0, rad);
        if (rad) p.arc(rad, rad, 180, 90);
      } else if (Math.abs(dy) >= Math.abs(dx)) {
        const [s, e] = base(w * (tx < w / 2 ? 0.2 : 0.5), w * (tx < w / 2 ? 0.35 : 0.65));
        p.m(rad, 0).l(s, 0).l(tx, ty).l(e, 0).l(w - rad, 0);
        if (rad) p.arc(rad, rad, 270, 90);
        p.l(w, h - rad);
        if (rad) p.arc(rad, rad, 0, 90);
        p.l(rad, h);
        if (rad) p.arc(rad, rad, 90, 90);
        p.l(0, rad);
        if (rad) p.arc(rad, rad, 180, 90);
      } else if (dx > 0) {
        const [s, e] = base(h * (ty < h / 2 ? 0.2 : 0.5), h * (ty < h / 2 ? 0.35 : 0.65));
        p.m(rad, 0).l(w - rad, 0);
        if (rad) p.arc(rad, rad, 270, 90);
        p.l(w, s).l(tx, ty).l(w, e).l(w, h - rad);
        if (rad) p.arc(rad, rad, 0, 90);
        p.l(rad, h);
        if (rad) p.arc(rad, rad, 90, 90);
        p.l(0, rad);
        if (rad) p.arc(rad, rad, 180, 90);
      } else {
        const [s, e] = base(h * (ty < h / 2 ? 0.2 : 0.5), h * (ty < h / 2 ? 0.35 : 0.65));
        p.m(rad, 0).l(w - rad, 0);
        if (rad) p.arc(rad, rad, 270, 90);
        p.l(w, h - rad);
        if (rad) p.arc(rad, rad, 0, 90);
        p.l(rad, h);
        if (rad) p.arc(rad, rad, 90, 90);
        p.l(0, e).l(tx, ty).l(0, s).l(0, rad);
        if (rad) p.arc(rad, rad, 180, 90);
      }
      return one(p.z().toString());
    }
    case 'cloud':
    case 'cloudCallout': {
      // bumpy outline made of arcs around an ellipse
      const n = 10;
      const pts = regularPolygon(n, w * 0.92, h * 0.88).map(([x, y]) => [x + w * 0.04, y + h * 0.06] as [number, number]);
      let d = `M${r(pts[0][0])} ${r(pts[0][1])}`;
      for (let i = 0; i < n; i++) {
        const [x2, y2] = pts[(i + 1) % n];
        const rr = Math.hypot(x2 - pts[i][0], y2 - pts[i][1]) * 0.62;
        d += `A${r(rr)} ${r(rr)} 0 0 1 ${r(x2)} ${r(y2)}`;
      }
      d += 'Z';
      return one(d, [w * 0.18, h * 0.2, w * 0.18, h * 0.2]);
    }
    case 'heart': {
      const d = P()
        .m(w / 2, h * 0.25)
        .c(w * 0.5, h * -0.05, w * -0.1, h * 0.02, w * 0.08, h * 0.42)
        .c(w * 0.2, h * 0.65, w * 0.42, h * 0.8, w / 2, h)
        .c(w * 0.58, h * 0.8, w * 0.8, h * 0.65, w * 0.92, h * 0.42)
        .c(w * 1.1, h * 0.02, w * 0.5, h * -0.05, w / 2, h * 0.25)
        .z()
        .toString();
      return one(d, [w * 0.2, h * 0.2, w * 0.2, h * 0.25]);
    }
    case 'lightningBolt': {
      const pts: [number, number][] = [
        [8458, 0],
        [0, 3923],
        [7564, 8416],
        [4993, 9720],
        [12197, 13904],
        [9987, 14934],
        [21600, 21600],
        [14768, 12911],
        [16558, 12016],
        [11030, 6840],
        [12831, 6120],
        [8458, 0],
      ];
      return one(poly(pts.map(([x, y]) => [(x / 21600) * w, (y / 21600) * h])), [w * 0.3, h * 0.3, w * 0.3, h * 0.3]);
    }
    case 'moon': {
      const t = w * a();
      const d = P().m(w, h).arc(w, h / 2, 90, 180).arc(w - t, h / 2, 270, -180).z().toString();
      return one(d);
    }
    case 'smileyFace': {
      const eye = Math.min(w, h) * 0.08;
      const mouth = P().m(w * 0.3, h * 0.65).q(w / 2, h * (0.65 + a() * 4), w * 0.7, h * 0.65).toString();
      return {
        paths: [
          { d: ellipsePath(0, 0, w, h), fill: 'norm', stroke: true },
          { d: ellipsePath(w * 0.32 - eye / 2, h * 0.35 - eye / 2, eye, eye) + ellipsePath(w * 0.68 - eye / 2, h * 0.35 - eye / 2, eye, eye), fill: 'darkenLess', stroke: true },
          { d: mouth, fill: 'none', stroke: true },
        ],
      };
    }
    case 'sun': {
      const rays: string[] = [];
      for (let i = 0; i < 8; i++) {
        const ang = (i * Math.PI) / 4;
        const c = Math.cos(ang);
        const s = Math.sin(ang);
        const tip = [w / 2 + (w / 2) * c, h / 2 + (h / 2) * s];
        const b1 = [w / 2 + w * 0.36 * Math.cos(ang - 0.18), h / 2 + h * 0.36 * Math.sin(ang - 0.18)];
        const b2 = [w / 2 + w * 0.36 * Math.cos(ang + 0.18), h / 2 + h * 0.36 * Math.sin(ang + 0.18)];
        rays.push(poly([tip as [number, number], b1 as [number, number], b2 as [number, number]]));
      }
      return { paths: [{ d: rays.join('') + ellipsePath(w * 0.25, h * 0.25, w * 0.5, h * 0.5), fill: 'norm', stroke: true }], text: [w * 0.3, h * 0.3, w * 0.3, h * 0.3] };
    }
    case 'teardrop': {
      const k = Math.min(2, a());
      const tx = w / 2 + (w / 2) * k;
      const ty = h / 2 - (h / 2) * k;
      const d = P().m(0, h / 2).arc(w / 2, h / 2, 180, 90).q(tx, 0 + (ty < 0 ? 0 : 0), Math.min(tx, w), Math.max(ty, 0)).q(w, 0, w, h / 2).arc(w / 2, h / 2, 0, 180).z().toString();
      return one(d, [w * 0.15, h * 0.15, w * 0.15, h * 0.15]);
    }
    case 'pie':
    case 'chord':
    case 'arc':
    case 'blockArc': {
      const st = (adj.adj1 ?? 0) / 60000;
      const en = (adj.adj2 ?? 16200000) / 60000;
      let sw = en - st;
      if (sw <= 0) sw += 360;
      const pt = (ang: number, rx = w / 2, ry = h / 2) => [w / 2 + rx * Math.cos((ang * Math.PI) / 180), h / 2 + ry * Math.sin((ang * Math.PI) / 180)];
      const [sx, sy] = pt(st);
      const arcP = P().m(sx, sy).arc(w / 2, h / 2, st, sw);
      if (geom === 'arc') return { paths: [{ d: arcP.toString(), fill: 'none', stroke: true }] };
      if (geom === 'chord') return one(arcP.z().toString());
      if (geom === 'pie') return one(arcP.l(w / 2, h / 2).z().toString());
      const t = Math.min(w, h) * ((adj.adj3 ?? 25000) / 100000);
      const [ex, ey] = pt(st + sw, w / 2 - t, h / 2 - t);
      const d = arcP.l(ex, ey).arc(w / 2 - t, h / 2 - t, st + sw, -sw).z().toString();
      return one(d);
    }
    case 'snip1Rect': {
      const x = ss * a();
      return one(poly([[0, 0], [w - x, 0], [w, x], [w, h], [0, h]]), [0, x / 2, x / 2, 0]);
    }
    case 'snip2SameRect': {
      const x = ss * a('adj1');
      const y = ss * a('adj2');
      return one(poly([[x, 0], [w - x, 0], [w, x], [w, h - y], [w - y, h], [y, h], [0, h - y], [0, x]]), [x / 2, x / 2, x / 2, y / 2]);
    }
    case 'round1Rect': {
      const x = Math.min(ss / 2, ss * a());
      return one(P().m(0, 0).l(w - x, 0).arc(x, x, 270, 90).l(w, h).l(0, h).z().toString());
    }
    case 'round2SameRect': {
      const x = Math.min(ss / 2, ss * a('adj1'));
      const y = Math.min(ss / 2, ss * a('adj2'));
      const p = P().m(x, 0).l(w - x, 0).arc(x, x, 270, 90).l(w, h - y);
      if (y) p.arc(y, y, 0, 90);
      p.l(y, h);
      if (y) p.arc(y, y, 90, 90);
      p.l(0, x).arc(x, x, 180, 90).z();
      return one(p.toString());
    }
    case 'plaque': {
      const x = ss * a();
      return one(P().m(0, x).arc(x, x, 90, -90).l(w - x, 0).arc(x, x, 180, -90).l(w, h - x).arc(x, x, 270, -90).l(x, h).arc(x, x, 0, -90).z().toString(), [x * 0.7, x * 0.7, x * 0.7, x * 0.7]);
    }
    case 'foldedCorner': {
      const x = ss * a();
      return {
        paths: [
          { d: poly([[0, 0], [w, 0], [w, h - x], [w - x, h], [0, h]]), fill: 'norm', stroke: true },
          { d: poly([[w - x, h], [w - x * 0.8, h - x * 0.8], [w, h - x]]), fill: 'darkenLess', stroke: true },
        ],
      };
    }
    case 'halfFrame': {
      const x = Math.min(w, ss * a('adj1'));
      const y = Math.min(h, ss * a('adj2'));
      return one(poly([[0, 0], [w, 0], [w - (w * x) / Math.max(1, h), x], [y, x], [y, h - (h * y) / Math.max(1, w)], [0, h]]));
    }
    case 'corner': {
      const x = ss * a('adj1');
      const y = ss * a('adj2');
      return one(poly([[0, 0], [y, 0], [y, h - x], [w, h - x], [w, h], [0, h]]));
    }
    case 'diagStripe': {
      const x = w * a();
      const y = h * a();
      return one(poly([[0, y], [x, 0], [w, 0], [0, h]]));
    }
    case 'flowChartTerminator':
      return one(roundRectPath(0, 0, w, h, Math.min(w, h) / 2), [h / 4, 0, h / 4, 0]);
    case 'flowChartDocument': {
      const d = P()
        .m(0, 0)
        .l(w, 0)
        .l(w, h * 0.83)
        .c(w * 0.75, h * 0.72, w * 0.5, h * 0.95, w * 0.25, h * 1.02)
        .q(w * 0.1, h * 1.05, 0, h * 0.95)
        .z()
        .toString();
      return one(d, [0, 0, 0, h * 0.17]);
    }
    case 'flowChartPredefinedProcess':
      return {
        paths: [
          { d: poly([[0, 0], [w, 0], [w, h], [0, h]]), fill: 'norm', stroke: true },
          { d: `M${r(w / 8)} 0L${r(w / 8)} ${r(h)}M${r((w * 7) / 8)} 0L${r((w * 7) / 8)} ${r(h)}`, fill: 'none', stroke: true },
        ],
        text: [w / 8, 0, w / 8, 0],
      };
    case 'flowChartManualInput':
      return one(poly([[0, h / 5], [w, 0], [w, h], [0, h]]), [0, h / 5, 0, 0]);
    case 'flowChartDelay':
      return one(P().m(0, 0).l(w / 2, 0).arc(w / 2, h / 2, 270, 180).l(0, h).z().toString());
    case 'flowChartOffpageConnector':
      return one(poly([[0, 0], [w, 0], [w, h * 0.8], [w / 2, h], [0, h * 0.8]]), [0, 0, 0, h * 0.2]);
    case 'flowChartSort':
      return { paths: [{ d: poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]), fill: 'norm', stroke: true }, { d: `M0 ${r(h / 2)}L${r(w)} ${r(h / 2)}`, fill: 'none', stroke: true }] };
    case 'wave':
    case 'doubleWave': {
      const amp = h * a('adj1');
      const d = P()
        .m(0, amp)
        .c(w / 3, -amp, (w * 2) / 3, amp * 3, w, amp)
        .l(w, h - amp)
        .c((w * 2) / 3, h + amp, w / 3, h - amp * 3, 0, h - amp)
        .z()
        .toString();
      return one(d, [0, amp * 2, 0, amp * 2]);
    }
    case 'mathPlus': {
      const t = ss * a('adj1');
      const x1 = (w - t) / 2;
      const y1 = (h - t) / 2;
      const m = ss * 0.12;
      return one(poly([[x1, m], [w - x1, m], [w - x1, y1], [w - m, y1], [w - m, h - y1], [w - x1, h - y1], [w - x1, h - m], [x1, h - m], [x1, h - y1], [m, h - y1], [m, y1], [x1, y1]]));
    }
    case 'mathMinus': {
      const t = h * a('adj1');
      return one(poly([[w * 0.12, (h - t) / 2], [w * 0.88, (h - t) / 2], [w * 0.88, (h + t) / 2], [w * 0.12, (h + t) / 2]]));
    }
    case 'mathEqual': {
      const t = h * a('adj1');
      const g = h * a('adj2');
      const y1 = h / 2 - g / 2 - t;
      const y2 = h / 2 + g / 2;
      return one(poly([[w * 0.12, y1], [w * 0.88, y1], [w * 0.88, y1 + t], [w * 0.12, y1 + t]]) + poly([[w * 0.12, y2], [w * 0.88, y2], [w * 0.88, y2 + t], [w * 0.12, y2 + t]]));
    }
    case 'mathMultiply': {
      const t = ss * a('adj1') * 0.7;
      const m = ss * 0.15;
      return one(poly([[m, m + t], [m + t, m], [w / 2, h / 2 - t], [w - m - t, m], [w - m, m + t], [w / 2 + t, h / 2], [w - m, h - m - t], [w - m - t, h - m], [w / 2, h / 2 + t], [m + t, h - m], [m, h - m - t], [w / 2 - t, h / 2]]));
    }
    case 'line':
    case 'straightConnector1':
      return { paths: [{ d: `M0 0L${r(w)} ${r(h)}`, fill: 'none', stroke: true }] };
    case 'bentConnector2':
      return { paths: [{ d: `M0 0L${r(w)} 0L${r(w)} ${r(h)}`, fill: 'none', stroke: true }] };
    case 'bentConnector3':
    case 'bentConnector4': {
      const x = w * a('adj1');
      return { paths: [{ d: `M0 0L${r(x)} 0L${r(x)} ${r(h)}L${r(w)} ${r(h)}`, fill: 'none', stroke: true }] };
    }
    case 'curvedConnector2':
      return { paths: [{ d: `M0 0Q${r(w)} 0 ${r(w)} ${r(h)}`, fill: 'none', stroke: true }] };
    case 'curvedConnector3':
    case 'curvedConnector4': {
      const x = w * a('adj1');
      return { paths: [{ d: `M0 0C${r(x)} 0 ${r(x)} ${r(h)} ${r(w)} ${r(h)}`, fill: 'none', stroke: true }] };
    }
    case 'bracketPair':
    case 'bracePair': {
      const x = ss * a();
      if (geom === 'bracketPair') {
        const d = P().m(x, h).arc(x, x, 90, 90).l(0, x).arc(x, x, 180, 90).toString() + P().m(w - x, 0).arc(x, x, 270, 90).l(w, h - x).arc(x, x, 0, 90).toString();
        return { paths: [{ d, fill: 'none', stroke: true }], text: [x, x, x, x] };
      }
      const d =
        P().m(x * 2, h).arc(x, x, 90, 90).l(x, h / 2 + x).arc(x, x, 0, -90).arc(x, x, 90, -90).l(x, x).arc(x, x, 180, 90).toString() +
        P().m(w - x * 2, 0).arc(x, x, 270, 90).l(w - x, h / 2 - x).arc(x, x, 180, -90).arc(x, x, 270, -90).l(w - x, h - x).arc(x, x, 0, 90).toString();
      return { paths: [{ d, fill: 'none', stroke: true }], text: [x * 2, x, x * 2, x] };
    }
    case 'leftBracket':
    case 'rightBracket': {
      const y = Math.min(h / 2, ss * a());
      const d = P().m(w, h).arc(w, y, 90, 90).l(0, y).arc(w, y, 180, 90).toString();
      return { paths: [{ d: geom === 'leftBracket' ? d : flipX(d, w), fill: 'none', stroke: true }] };
    }
    case 'leftBrace':
    case 'rightBrace': {
      const y = Math.min(h / 4, ss * a('adj1'));
      const mid = h * a('adj2');
      const hw = w / 2;
      const d = P().m(w, h).arc(hw, y, 90, 90).l(hw, mid + y).arc(hw, y, 0, -90).arc(hw, y, 90, -90).l(hw, y).arc(hw, y, 180, 90).toString();
      return { paths: [{ d: geom === 'leftBrace' ? d : flipX(d, w), fill: 'none', stroke: true }] };
    }
    default:
      return one(poly([[0, 0], [w, 0], [w, h], [0, h]]));
  }
}

const RECT_PROBE = presetGeometry('rect', 100, 60).paths[0].d;
const known = new Map<string, boolean>();

/** Whether a preset has geometry of its own here (unknown presets draw as plain rectangles). */
export function isKnownPreset(geom: string): boolean {
  if (geom === 'rect' || geom === 'textBox' || geom === 'flowChartProcess') return true;
  let hit = known.get(geom);
  if (hit === undefined) {
    const g = presetGeometry(geom, 100, 60);
    hit = g.paths.length !== 1 || g.paths[0].d !== RECT_PROBE;
    known.set(geom, hit);
  }
  return hit;
}

function flipX(d: string, w: number): string {
  // mirrors absolute path coordinates horizontally (M/L/A/C/Q with absolute numbers)
  return d.replace(/([MLCQ])([^MLCQAZ]*)|A([^MLCQAZ]*)/g, (_m, cmd: string | undefined, args: string | undefined, arcArgs: string | undefined) => {
    if (arcArgs !== undefined) {
      const n = arcArgs.trim().split(/[\s,]+/).map(Number);
      return `A${n[0]} ${n[1]} ${n[2]} ${n[3]} ${n[4] ? 0 : 1} ${r(w - n[5])} ${n[6]}`;
    }
    const n = (args ?? '').trim().split(/[\s,]+/).filter(Boolean).map(Number);
    for (let i = 0; i < n.length; i += 2) n[i] = r(w - n[i]);
    return `${cmd}${n.join(' ')}`;
  });
}

/** Custom geometry (a:custGeom) scaled to the shape's box. */
export function customGeometry(paths: CustomPath[], w: number, h: number): Geometry {
  return {
    paths: paths.map((p) => {
      const sx = p.w ? w / p.w : 1;
      const sy = p.h ? h / p.h : 1;
      const b = P();
      for (const cmd of p.cmds) {
        const q = cmd.p;
        switch (cmd.c) {
          case 'M':
            b.m(q[0] * sx, q[1] * sy);
            break;
          case 'L':
            b.l(q[0] * sx, q[1] * sy);
            break;
          case 'C':
            b.c(q[0] * sx, q[1] * sy, q[2] * sx, q[3] * sy, q[4] * sx, q[5] * sy);
            break;
          case 'Q':
            b.q(q[0] * sx, q[1] * sy, q[2] * sx, q[3] * sy);
            break;
          case 'A':
            b.arc(q[0] * sx, q[1] * sy, q[2], q[3]);
            break;
          case 'Z':
            b.z();
            break;
        }
      }
      return { d: b.toString(), fill: p.fill === false ? 'none' : 'norm', stroke: p.stroke !== false };
    }),
  };
}

/* ------------------------------------------------------------- gallery */

export interface ShapeGroup {
  label: string;
  shapes: { geom: string; label: string }[];
}

export const SHAPE_GALLERY: ShapeGroup[] = [
  {
    label: 'Lines',
    shapes: [
      { geom: 'line', label: 'Line' },
      { geom: 'straightConnector1:arrow', label: 'Arrow' },
      { geom: 'straightConnector1:double', label: 'Double arrow' },
      { geom: 'bentConnector3', label: 'Elbow connector' },
      { geom: 'curvedConnector3', label: 'Curved connector' },
    ],
  },
  {
    label: 'Rectangles',
    shapes: [
      { geom: 'rect', label: 'Rectangle' },
      { geom: 'roundRect', label: 'Rounded rectangle' },
      { geom: 'snip1Rect', label: 'Snipped corner' },
      { geom: 'snip2SameRect', label: 'Snipped same side' },
      { geom: 'round1Rect', label: 'Round one corner' },
      { geom: 'round2SameRect', label: 'Round same side' },
      { geom: 'plaque', label: 'Plaque' },
      { geom: 'foldedCorner', label: 'Folded corner' },
    ],
  },
  {
    label: 'Basic shapes',
    shapes: [
      { geom: 'ellipse', label: 'Oval' },
      { geom: 'triangle', label: 'Triangle' },
      { geom: 'rtTriangle', label: 'Right triangle' },
      { geom: 'parallelogram', label: 'Parallelogram' },
      { geom: 'trapezoid', label: 'Trapezoid' },
      { geom: 'diamond', label: 'Diamond' },
      { geom: 'pentagon', label: 'Pentagon' },
      { geom: 'hexagon', label: 'Hexagon' },
      { geom: 'heptagon', label: 'Heptagon' },
      { geom: 'octagon', label: 'Octagon' },
      { geom: 'decagon', label: 'Decagon' },
      { geom: 'dodecagon', label: 'Dodecagon' },
      { geom: 'pie', label: 'Partial circle' },
      { geom: 'chord', label: 'Chord' },
      { geom: 'teardrop', label: 'Teardrop' },
      { geom: 'frame', label: 'Frame' },
      { geom: 'halfFrame', label: 'Half frame' },
      { geom: 'corner', label: 'L-shape' },
      { geom: 'diagStripe', label: 'Diagonal stripe' },
      { geom: 'plus', label: 'Cross' },
      { geom: 'can', label: 'Cylinder' },
      { geom: 'cube', label: 'Cube' },
      { geom: 'bevel', label: 'Bevel' },
      { geom: 'donut', label: 'Donut' },
      { geom: 'blockArc', label: 'Block arc' },
      { geom: 'smileyFace', label: 'Smiley face' },
      { geom: 'heart', label: 'Heart' },
      { geom: 'lightningBolt', label: 'Lightning bolt' },
      { geom: 'sun', label: 'Sun' },
      { geom: 'moon', label: 'Moon' },
      { geom: 'cloud', label: 'Cloud' },
      { geom: 'arc', label: 'Arc' },
      { geom: 'bracketPair', label: 'Double bracket' },
      { geom: 'bracePair', label: 'Double brace' },
      { geom: 'leftBracket', label: 'Left bracket' },
      { geom: 'rightBracket', label: 'Right bracket' },
      { geom: 'leftBrace', label: 'Left brace' },
      { geom: 'rightBrace', label: 'Right brace' },
    ],
  },
  {
    label: 'Block arrows',
    shapes: [
      { geom: 'rightArrow', label: 'Right arrow' },
      { geom: 'leftArrow', label: 'Left arrow' },
      { geom: 'upArrow', label: 'Up arrow' },
      { geom: 'downArrow', label: 'Down arrow' },
      { geom: 'leftRightArrow', label: 'Left-right arrow' },
      { geom: 'upDownArrow', label: 'Up-down arrow' },
      { geom: 'notchedRightArrow', label: 'Notched arrow' },
      { geom: 'homePlate', label: 'Pentagon arrow' },
      { geom: 'chevron', label: 'Chevron' },
    ],
  },
  {
    label: 'Equation shapes',
    shapes: [
      { geom: 'mathPlus', label: 'Plus' },
      { geom: 'mathMinus', label: 'Minus' },
      { geom: 'mathMultiply', label: 'Multiply' },
      { geom: 'mathEqual', label: 'Equal' },
    ],
  },
  {
    label: 'Flowchart',
    shapes: [
      { geom: 'flowChartProcess', label: 'Process' },
      { geom: 'flowChartAlternateProcess', label: 'Alternate process' },
      { geom: 'flowChartDecision', label: 'Decision' },
      { geom: 'flowChartInputOutput', label: 'Data' },
      { geom: 'flowChartPredefinedProcess', label: 'Predefined process' },
      { geom: 'flowChartDocument', label: 'Document' },
      { geom: 'flowChartTerminator', label: 'Terminator' },
      { geom: 'flowChartPreparation', label: 'Preparation' },
      { geom: 'flowChartManualInput', label: 'Manual input' },
      { geom: 'flowChartManualOperation', label: 'Manual operation' },
      { geom: 'flowChartConnector', label: 'Connector' },
      { geom: 'flowChartOffpageConnector', label: 'Off-page connector' },
      { geom: 'flowChartDelay', label: 'Delay' },
      { geom: 'flowChartMagneticDisk', label: 'Stored data' },
      { geom: 'flowChartMerge', label: 'Merge' },
      { geom: 'flowChartExtract', label: 'Extract' },
      { geom: 'flowChartSort', label: 'Sort' },
      { geom: 'flowChartOr', label: 'Or' },
    ],
  },
  {
    label: 'Stars and banners',
    shapes: [
      { geom: 'star4', label: '4-point star' },
      { geom: 'star5', label: '5-point star' },
      { geom: 'star6', label: '6-point star' },
      { geom: 'star7', label: '7-point star' },
      { geom: 'star8', label: '8-point star' },
      { geom: 'star10', label: '10-point star' },
      { geom: 'star12', label: '12-point star' },
      { geom: 'star16', label: '16-point star' },
      { geom: 'star24', label: '24-point star' },
      { geom: 'wave', label: 'Wave' },
      { geom: 'doubleWave', label: 'Double wave' },
    ],
  },
  {
    label: 'Callouts',
    shapes: [
      { geom: 'wedgeRectCallout', label: 'Rectangular callout' },
      { geom: 'wedgeRoundRectCallout', label: 'Rounded callout' },
      { geom: 'wedgeEllipseCallout', label: 'Oval callout' },
      { geom: 'cloudCallout', label: 'Cloud callout' },
    ],
  },
];

/** Adjust handle (the yellow diamond): where it sits for the current values and what values a drag produces. */
export interface AdjHandle {
  pos: (w: number, h: number, adj: Record<string, number>) => [number, number];
  apply: (w: number, h: number, x: number, y: number) => Record<string, number>;
}

const clampV = (v: number, lo: number, hi: number) => Math.round(Math.max(lo, Math.min(hi, v)));

export function adjHandles(geom: string): AdjHandle[] {
  const val = (adj: Record<string, number>, k: string) => adj[k] ?? DEFAULT_ADJ[geom]?.[k] ?? 0;
  const fromLeft = (key: string, max = 50000): AdjHandle => ({
    pos: (w, h, adj) => [(Math.min(w, h) * val(adj, key)) / 100000, 0],
    apply: (w, h, x) => ({ [key]: clampV((x / Math.min(w, h)) * 100000, 0, max) }),
  });
  const fromRight = (key: string, max: (w: number, h: number) => number): AdjHandle => ({
    pos: (w, h, adj) => [w - (Math.min(w, h) * val(adj, key)) / 100000, 0],
    apply: (w, h, x) => ({ [key]: clampV(((w - x) / Math.min(w, h)) * 100000, 0, max(w, h)) }),
  });
  switch (geom) {
    case 'roundRect':
    case 'octagon':
    case 'plus':
    case 'donut':
    case 'can':
    case 'cube':
    case 'bevel':
    case 'plaque':
    case 'foldedCorner':
    case 'round1Rect':
      return [fromLeft('adj')];
    case 'frame':
      return [fromLeft('adj1')];
    case 'parallelogram':
    case 'trapezoid':
    case 'hexagon':
      return [fromLeft('adj', 100000)];
    case 'triangle':
      return [{ pos: (w, _h, adj) => [(w * val(adj, 'adj')) / 100000, 0], apply: (w, _h, x) => ({ adj: clampV((x / w) * 100000, 0, 100000) }) }];
    case 'chevron':
    case 'homePlate':
    case 'snip1Rect':
      return [fromRight('adj', (w, h) => (w / Math.min(w, h)) * 100000)];
    case 'rightArrow':
      return [fromRight('adj2', (w, h) => (w / Math.min(w, h)) * 100000)];
    case 'leftArrow':
      return [fromLeft('adj2', 100000)];
    case 'wedgeRectCallout':
    case 'wedgeRoundRectCallout':
    case 'wedgeEllipseCallout':
    case 'cloudCallout':
      return [
        {
          pos: (w, h, adj) => [w / 2 + (w * val(adj, 'adj1')) / 100000, h / 2 + (h * val(adj, 'adj2')) / 100000],
          apply: (w, h, x, y) => ({ adj1: Math.round(((x - w / 2) / w) * 100000), adj2: Math.round(((y - h / 2) / h) * 100000) }),
        },
      ];
    case 'star4':
    case 'star5':
    case 'star6':
    case 'star7':
    case 'star8':
    case 'star10':
    case 'star12':
    case 'star16':
    case 'star24':
      return [{ pos: (w, h, adj) => [w / 2, h / 2 - (h / 2) * ((val(adj, 'adj') * 2) / 100000)], apply: (_w, h, _x, y) => ({ adj: clampV(((h / 2 - y) / (h / 2)) * 50000, 1000, 50000) }) }];
    default:
      return [];
  }
}
