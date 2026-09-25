/**
 * Shape geometry for OpenDocument presentations.
 *
 * Writing: our preset shapes are drawn as SVG path data; ODF custom shapes take a `draw:enhanced-path`,
 * which has no SVG-style arcs, so arcs become cubic Béziers. Reading: an ODF custom shape describes its
 * outline with equations, modifiers and path commands (including the angle and quadrant arcs of the
 * old binary formats and LibreOffice's OOXML-style `G` arc); these are evaluated into plain paths.
 */
import type { GeomPath } from '../geometry';
import type { CustomPath, PathCmd } from '../model';

/* ================================================================ writing */

type Seg = { c: 'M' | 'L' | 'C' | 'Z'; p: number[] };

/** Parses SVG path data (absolute or relative, with smooth curves and arcs) into moves, lines, cubics and closes. */
export function svgPathSegments(d: string): Seg[] {
  const tokens = d.match(/[MLHVCSQTAZmlhvcsqtaz]|[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g) ?? [];
  const out: Seg[] = [];
  let i = 0;
  let cmd = '';
  let x = 0;
  let y = 0;
  let sx = 0;
  let sy = 0;
  // the last control point, for smooth curves
  let cx: number | null = null;
  let cy: number | null = null;
  let qx: number | null = null;
  let qy: number | null = null;
  const num = () => Number(tokens[i++]);
  const quad = (x1: number, y1: number, ex: number, ey: number) => {
    // a quadratic is an exact cubic with control points two thirds of the way to the quadratic one
    out.push({ c: 'C', p: [x + ((x1 - x) * 2) / 3, y + ((y1 - y) * 2) / 3, ex + ((x1 - ex) * 2) / 3, ey + ((y1 - ey) * 2) / 3, ex, ey] });
  };
  while (i < tokens.length) {
    const t = tokens[i];
    if (/^[a-zA-Z]$/.test(t)) {
      cmd = t;
      i++;
      if (cmd === 'Z' || cmd === 'z') {
        out.push({ c: 'Z', p: [] });
        x = sx;
        y = sy;
        cx = cy = qx = qy = null;
        continue;
      }
    } else if (!cmd) {
      i++;
      continue;
    }
    const rel = cmd === cmd.toLowerCase();
    const ox = rel ? x : 0;
    const oy = rel ? y : 0;
    const before = i;
    switch (cmd.toUpperCase()) {
      case 'M':
        x = ox + num();
        y = oy + num();
        sx = x;
        sy = y;
        out.push({ c: 'M', p: [x, y] });
        cmd = rel ? 'l' : 'L'; // further pairs are line-tos
        cx = cy = qx = qy = null;
        break;
      case 'L':
        x = ox + num();
        y = oy + num();
        out.push({ c: 'L', p: [x, y] });
        cx = cy = qx = qy = null;
        break;
      case 'H':
        x = ox + num();
        out.push({ c: 'L', p: [x, y] });
        cx = cy = qx = qy = null;
        break;
      case 'V':
        y = oy + num();
        out.push({ c: 'L', p: [x, y] });
        cx = cy = qx = qy = null;
        break;
      case 'C':
      case 'S': {
        let x1: number;
        let y1: number;
        if (cmd.toUpperCase() === 'C') {
          x1 = ox + num();
          y1 = oy + num();
        } else {
          x1 = cx === null ? x : 2 * x - cx;
          y1 = cy === null ? y : 2 * y - cy;
        }
        const x2 = ox + num();
        const y2 = oy + num();
        const ex = ox + num();
        const ey = oy + num();
        out.push({ c: 'C', p: [x1, y1, x2, y2, ex, ey] });
        cx = x2;
        cy = y2;
        qx = qy = null;
        x = ex;
        y = ey;
        break;
      }
      case 'Q':
      case 'T': {
        let x1: number;
        let y1: number;
        if (cmd.toUpperCase() === 'Q') {
          x1 = ox + num();
          y1 = oy + num();
        } else {
          x1 = qx === null ? x : 2 * x - qx;
          y1 = qy === null ? y : 2 * y - qy;
        }
        const ex = ox + num();
        const ey = oy + num();
        quad(x1, y1, ex, ey);
        qx = x1;
        qy = y1;
        cx = cy = null;
        x = ex;
        y = ey;
        break;
      }
      case 'A': {
        const [rx, ry, rot, large, sweep] = [num(), num(), num(), num(), num()];
        const ex = ox + num();
        const ey = oy + num();
        for (const c of arcToCubics(x, y, rx, ry, rot, !!large, !!sweep, ex, ey)) out.push({ c: 'C', p: c });
        x = ex;
        y = ey;
        cx = cy = qx = qy = null;
        break;
      }
      default:
        i++;
    }
    if (i === before) i++;
  }
  return out.filter((s) => s.p.every((v) => Number.isFinite(v)));
}

/** Converts an SVG elliptical arc to cubic Béziers (at most 90° each). */
export function arcToCubics(x1: number, y1: number, rx: number, ry: number, rotDeg: number, large: boolean, sweep: boolean, x2: number, y2: number): number[][] {
  if (!rx || !ry || (x1 === x2 && y1 === y2)) return x1 === x2 && y1 === y2 ? [] : [[x1, y1, x2, y2, x2, y2]];
  rx = Math.abs(rx);
  ry = Math.abs(ry);
  const phi = (rotDeg * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  // endpoint to centre parameterisation (SVG implementation notes, F.6.5)
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy;
  const y1p = -sin * dx + cos * dy;
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let k = Math.sqrt(Math.max(0, num / den));
  if (large === sweep) k = -k;
  const cxp = (k * rx * y1p) / ry;
  const cyp = (-k * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const angle = (ux: number, uy: number, vx: number, vy: number) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const t1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dt = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!sweep && dt > 0) dt -= 2 * Math.PI;
  else if (sweep && dt < 0) dt += 2 * Math.PI;
  return ellipseCubics(cx, cy, rx, ry, phi, t1, dt);
}

/** Cubic Béziers along an ellipse from parameter angle t1 through dt (radians, y-down). */
function ellipseCubics(cx: number, cy: number, rx: number, ry: number, phi: number, t1: number, dt: number): number[][] {
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
  const step = dt / n;
  const kappa = (4 / 3) * Math.tan(step / 4);
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const pt = (t: number) => {
    const ex = rx * Math.cos(t);
    const ey = ry * Math.sin(t);
    return [cx + cos * ex - sin * ey, cy + sin * ex + cos * ey];
  };
  const der = (t: number) => {
    const ex = -rx * Math.sin(t);
    const ey = ry * Math.cos(t);
    return [cos * ex - sin * ey, sin * ex + cos * ey];
  };
  const out: number[][] = [];
  let a = t1;
  for (let i = 0; i < n; i++) {
    const b = a + step;
    const p0 = pt(a);
    const p1 = pt(b);
    const d0 = der(a);
    const d1 = der(b);
    out.push([p0[0] + kappa * d0[0], p0[1] + kappa * d0[1], p1[0] - kappa * d1[0], p1[1] - kappa * d1[1], p1[0], p1[1]]);
    a = b;
  }
  return out;
}

const FILL_CMD: Record<string, string> = { darken: 'H', darkenLess: 'I', lighten: 'J', lightenLess: 'K' };

export interface EnhancedPathOut {
  /** Standard ODF path (lines and curves only). */
  path: string;
  /** LibreOffice's extended path with shading commands, when a subpath is darkened or lightened. */
  extended?: string;
  viewBox: string;
  textAreas?: string;
  /** draw:equation elements the text areas refer to. */
  equations?: string;
}

/** Hundredths of a millimetre per pixel (ODF's logical shape units). */
const LOG = 2540 / 96;

/**
 * ODF enhanced geometry for a shape drawn at w × h px. The view box is the shape's own size in
 * hundredths of a millimetre, so coordinates mean the same whether or not an app scales them.
 * LibreOffice reads the text areas of its OOXML preset types ("ooxml-…") without scaling, so for
 * those they are written as fractions of the shape's logical size, which stays right after resizing.
 */
export function enhancedPath(paths: GeomPath[], w: number, h: number, text?: [number, number, number, number], logicalText = false): EnhancedPathOut {
  const q = (v: number) => String(Math.round(v * LOG));
  const parts: string[] = [];
  const ext: string[] = [];
  let shaded = false;
  for (const p of paths) {
    const segs = svgPathSegments(p.d);
    if (!segs.length) continue;
    let body = '';
    for (const s of segs) {
      if (s.c === 'Z') body += ' Z';
      else body += ` ${s.c} ${s.p.map(q).join(' ')}`;
    }
    const flags = `${p.fill === 'none' ? 'F ' : ''}${p.stroke ? '' : 'S '}`;
    parts.push(`${flags}${body.trim()} N`);
    const shade = FILL_CMD[p.fill];
    if (shade) shaded = true;
    ext.push(`${shade ? `${shade} ` : ''}${flags}${body.trim()} N`);
  }
  const W = Math.max(1, Math.round(w * LOG));
  const H = Math.max(1, Math.round(h * LOG));
  const out: EnhancedPathOut = { path: parts.join(' '), viewBox: `0 0 ${W} ${H}` };
  if (shaded) out.extended = ext.join(' ');
  if (text && text.some((v) => Math.abs(v) > 0.01)) {
    const box = [text[0], text[1], w - text[2], h - text[3]];
    if (logicalText) {
      const frac = (v: number, d: number) => (d > 0 ? Math.round((v / d) * 1e6) / 1e6 : 0);
      const f = [`logwidth*${frac(box[0], w)}`, `logheight*${frac(box[1], h)}`, `logwidth*${frac(box[2], w)}`, `logheight*${frac(box[3], h)}`];
      out.equations = f.map((formula, i) => `<draw:equation draw:name="affice${i}" draw:formula="${formula}"/>`).join('');
      out.textAreas = '?affice0 ?affice1 ?affice2 ?affice3';
    } else out.textAreas = box.map(q).join(' ');
  }
  return out;
}

/* ================================================================ reading */

export interface EnhancedGeometryIn {
  /** svg:viewBox as x, y, w, h (absent or all zero: the shape's own size in 1/100 mm). */
  viewBox?: number[];
  /** draw:enhanced-path, or drawooo:enhanced-path when present (it keeps arcs the standard path drops). */
  path: string;
  equations: Map<string, string>;
  modifiers: number[];
  textAreas?: string;
  /** Text areas in the shape's logical units (1/100 mm) rather than view-box units, as LibreOffice reads them for its OOXML shape types. */
  logicalText?: boolean;
}

export interface EvaluatedGeometry {
  paths: CustomPath[];
  /** Text rectangle as insets l, t, r, b in px from the shape's edges. */
  text?: [number, number, number, number];
}

type Fn = () => number;

/** Evaluates ODF draw:formula expressions (with ?fN equations and $N modifiers). */
class Formulas {
  private cache = new Map<string, number>();
  private busy = new Set<string>();
  constructor(
    private eqs: Map<string, string>,
    private mods: number[],
    private vars: Record<string, number>,
  ) {}

  equation(name: string): number {
    const hit = this.cache.get(name);
    if (hit !== undefined) return hit;
    const f = this.eqs.get(name);
    if (f === undefined || this.busy.has(name)) return 0;
    this.busy.add(name);
    let v = 0;
    try {
      v = this.compile(f)();
    } catch {
      v = 0;
    }
    this.busy.delete(name);
    if (!Number.isFinite(v)) v = 0;
    this.cache.set(name, v);
    return v;
  }

  /** A path parameter: a number, ?fN, $N or a bare identifier. */
  param(tok: string): number {
    if (tok[0] === '?') return this.equation(tok.slice(1));
    if (tok[0] === '$') return this.mods[Number(tok.slice(1))] ?? 0;
    const n = Number(tok);
    if (Number.isFinite(n)) return n;
    return this.vars[tok.toLowerCase()] ?? 0;
  }

  compile(src: string): Fn {
    const toks = src.match(/\?[A-Za-z0-9_]+|\$\d+|[A-Za-z_][A-Za-z0-9_]*|(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?|[-+*/(),]/g) ?? [];
    let i = 0;
    const peek = () => toks[i];
    const next = () => toks[i++];
    const expr = (): Fn => {
      let a = term();
      while (peek() === '+' || peek() === '-') {
        const op = next();
        const l = a;
        const r = term();
        a = op === '+' ? () => l() + r() : () => l() - r();
      }
      return a;
    };
    const term = (): Fn => {
      let a = unary();
      while (peek() === '*' || peek() === '/') {
        const op = next();
        const l = a;
        const r = unary();
        a = op === '*' ? () => l() * r() : () => {
          const d = r();
          return d === 0 ? 0 : l() / d;
        };
      }
      return a;
    };
    const unary = (): Fn => {
      if (peek() === '-') {
        next();
        const a = unary();
        return () => -a();
      }
      if (peek() === '+') {
        next();
        return unary();
      }
      return primary();
    };
    const args = (): Fn[] => {
      const out: Fn[] = [];
      if (next() !== '(') throw new Error('(');
      if (peek() === ')') {
        next();
        return out;
      }
      for (;;) {
        out.push(expr());
        const t = next();
        if (t === ')') break;
        if (t !== ',') throw new Error(',');
      }
      return out;
    };
    const primary = (): Fn => {
      const t = next();
      if (t === undefined) throw new Error('end');
      if (t === '(') {
        const a = expr();
        if (next() !== ')') throw new Error(')');
        return a;
      }
      if (t[0] === '?') {
        const name = t.slice(1);
        return () => this.equation(name);
      }
      if (t[0] === '$') {
        const k = Number(t.slice(1));
        return () => this.mods[k] ?? 0;
      }
      if (/^[\d.]/.test(t)) {
        const v = Number(t);
        return () => v;
      }
      const id = t.toLowerCase();
      if (peek() === '(') {
        const a = args();
        const f = (k: number) => a[k] ?? (() => 0);
        switch (id) {
          case 'abs':
            return () => Math.abs(f(0)());
          case 'sqrt':
            return () => Math.sqrt(Math.max(0, f(0)()));
          case 'sin':
            return () => Math.sin(f(0)());
          case 'cos':
            return () => Math.cos(f(0)());
          case 'tan':
            return () => Math.tan(f(0)());
          case 'atan':
            return () => Math.atan(f(0)());
          case 'atan2':
            // ODF: atan2(y, x), like C's atan2 (LibreOffice writes OOXML "at2 x y" as atan2(y,x))
            return () => Math.atan2(f(0)(), f(1)());
          case 'min':
            return () => Math.min(f(0)(), f(1)());
          case 'max':
            return () => Math.max(f(0)(), f(1)());
          case 'if':
            return () => (f(0)() > 0 ? f(1)() : f(2)());
          default:
            return () => 0;
        }
      }
      if (id === 'pi') return () => Math.PI;
      return () => this.vars[id] ?? 0;
    };
    const fn = expr();
    return fn;
  }
}

/** Splits a path into command letters and parameter tokens. */
function pathTokens(path: string): string[] {
  return path.match(/[A-Za-z](?![A-Za-z0-9_])|\?[A-Za-z0-9_]+|\$\d+|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?|[A-Za-z_][A-Za-z0-9_]+/g) ?? [];
}

const PARAMS: Record<string, number> = { M: 2, L: 2, C: 6, Q: 4, Z: 0, N: 0, F: 0, S: 0, T: 6, U: 6, A: 8, B: 8, W: 8, V: 8, X: 2, Y: 2, G: 4, H: 0, I: 0, J: 0, K: 0 };

/**
 * Evaluates enhanced geometry for a shape of w × h px into paths (in view-box units, arcs as cubic curves)
 * and the text rectangle.
 */
export function evaluateEnhancedGeometry(g: EnhancedGeometryIn, w: number, h: number): EvaluatedGeometry {
  const logW = (w * 2540) / 96;
  const logH = (h * 2540) / 96;
  const vb = g.viewBox && g.viewBox.length === 4 && (g.viewBox[2] > 0 || g.viewBox[3] > 0) ? g.viewBox : [0, 0, logW, logH];
  const [vx, vy] = vb;
  const vw = vb[2] || logW;
  const vh = vb[3] || logH;
  const vars: Record<string, number> = {
    left: vx,
    top: vy,
    right: vx + vw,
    bottom: vy + vh,
    width: vw,
    height: vh,
    logwidth: logW,
    logheight: logH,
    xstretch: 0,
    ystretch: 0,
    hasstroke: 1,
    hasfill: 1,
  };
  const f = new Formulas(g.equations, g.modifiers, vars);
  const toks = pathTokens(g.path);
  const paths: CustomPath[] = [];
  let cmds: PathCmd[] = [];
  let fill = true;
  let stroke = true;
  let cx = 0;
  let cy = 0;
  let open = false;
  const X = (v: number) => v - vx;
  const Y = (v: number) => v - vy;
  const move = (x: number, y: number) => {
    cmds.push({ c: 'M', p: [X(x), Y(y)] });
    cx = x;
    cy = y;
    open = true;
  };
  const line = (x: number, y: number) => {
    if (!open) return move(x, y);
    cmds.push({ c: 'L', p: [X(x), Y(y)] });
    cx = x;
    cy = y;
  };
  const cubics = (list: number[][]) => {
    for (const c of list) cmds.push({ c: 'C', p: [X(c[0]), Y(c[1]), X(c[2]), Y(c[3]), X(c[4]), Y(c[5])] });
    if (list.length) {
      const last = list[list.length - 1];
      cx = last[4];
      cy = last[5];
    }
  };
  const flush = () => {
    if (cmds.some((c) => c.c !== 'M')) paths.push({ w: vw, h: vh, cmds, ...(fill ? {} : { fill: false }), ...(stroke ? {} : { stroke: false }) });
    cmds = [];
    fill = true;
    stroke = true;
    open = false;
  };
  /** An arc of the ellipse inscribed in (x1,y1)-(x2,y2), from the ray through (x3,y3) to the ray through (x4,y4). */
  const arcRect = (x1: number, y1: number, x2: number, y2: number, x3: number, y3: number, x4: number, y4: number, clockwise: boolean, connect: boolean) => {
    const ecx = (x1 + x2) / 2;
    const ecy = (y1 + y2) / 2;
    const rx = Math.abs(x2 - x1) / 2;
    const ry = Math.abs(y2 - y1) / 2;
    if (!rx || !ry) return;
    const a1 = Math.atan2((y3 - ecy) / ry, (x3 - ecx) / rx);
    let a2 = Math.atan2((y4 - ecy) / ry, (x4 - ecx) / rx);
    // on screen (y down) "clockwise" means increasing angle
    if (clockwise) {
      while (a2 <= a1) a2 += 2 * Math.PI;
    } else {
      while (a2 >= a1) a2 -= 2 * Math.PI;
    }
    const sx = ecx + rx * Math.cos(a1);
    const sy = ecy + ry * Math.sin(a1);
    if (connect && open) line(sx, sy);
    else move(sx, sy);
    cubics(ellipseCubics(ecx, ecy, rx, ry, 0, a1, a2 - a1));
  };
  /** Angle arc: centre, radii and start/end angles in degrees, counter-clockwise with y pointing up. */
  const arcAngle = (x: number, y: number, rx: number, ry: number, t0: number, t1: number, connect: boolean) => {
    if (!rx || !ry) return;
    // degrees counter-clockwise (y up) become radians clockwise on screen: negate
    const a1 = (-t0 * Math.PI) / 180;
    let sweep = (-(t1 - t0) * Math.PI) / 180;
    // a sweep from 0 to 360 must go all the way round
    if (Math.abs(t1 - t0) >= 360) sweep = -2 * Math.PI;
    else if (sweep > 0) sweep -= 2 * Math.PI;
    const sx = x + rx * Math.cos(a1);
    const sy = y + ry * Math.sin(a1);
    if (connect && open) line(sx, sy);
    else move(sx, sy);
    cubics(ellipseCubics(x, y, rx, ry, 0, a1, sweep));
  };
  let i = 0;
  let cmd = '';
  let quadrantX = true;
  while (i < toks.length) {
    const t = toks[i];
    if (/^[A-Z]$/.test(t) && PARAMS[t] !== undefined) {
      cmd = t;
      i++;
      if (cmd === 'X' || cmd === 'Y') quadrantX = cmd === 'X';
      switch (cmd) {
        case 'Z':
          if (open) cmds.push({ c: 'Z', p: [] });
          open = false;
          continue;
        case 'N':
          flush();
          continue;
        case 'F':
          fill = false;
          continue;
        case 'S':
          stroke = false;
          continue;
        case 'H':
        case 'I':
        case 'J':
        case 'K':
          continue;
      }
      continue;
    }
    const n = PARAMS[cmd] ?? 0;
    if (!n || i + n > toks.length) {
      i++;
      continue;
    }
    const v = toks.slice(i, i + n).map((x) => f.param(x));
    i += n;
    switch (cmd) {
      case 'M':
        move(v[0], v[1]);
        cmd = 'L'; // further pairs after a move are line-tos
        break;
      case 'L':
        line(v[0], v[1]);
        break;
      case 'C':
        if (!open) move(cx, cy);
        cubics([v]);
        break;
      case 'Q':
        if (!open) move(cx, cy);
        cubics([[cx + ((v[0] - cx) * 2) / 3, cy + ((v[1] - cy) * 2) / 3, v[2] + ((v[0] - v[2]) * 2) / 3, v[3] + ((v[1] - v[3]) * 2) / 3, v[2], v[3]]]);
        break;
      case 'T':
        arcAngle(v[0], v[1], v[2], v[3], v[4], v[5], true);
        break;
      case 'U':
        arcAngle(v[0], v[1], v[2], v[3], v[4], v[5], false);
        break;
      case 'A':
        arcRect(v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], false, true);
        break;
      case 'B':
        arcRect(v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], false, false);
        break;
      case 'W':
        arcRect(v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], true, true);
        break;
      case 'V':
        arcRect(v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], true, false);
        break;
      case 'X':
      case 'Y': {
        // quarter ellipse to (x, y); X starts tangential to the x-axis, then the commands alternate
        if (!open) move(cx, cy);
        const [ex, ey] = v;
        const k = 0.5522847498;
        if (quadrantX) cubics([[cx + (ex - cx) * k, cy, ex, ey - (ey - cy) * k, ex, ey]]);
        else cubics([[cx, cy + (ey - cy) * k, ex - (ex - cx) * k, ey, ex, ey]]);
        quadrantX = !quadrantX;
        break;
      }
      case 'G': {
        // OOXML arcTo: radii, start angle and sweep in degrees, clockwise on screen
        const [wr, hr, st, sw] = v;
        if (!open) move(cx, cy);
        if (!wr || !hr || !sw) break;
        // OOXML angles are visual angles; convert them to the ellipse's parameter angle
        const param = (deg: number) => {
          const r = (deg * Math.PI) / 180;
          return Math.atan2(Math.sin(r) * wr, Math.cos(r) * hr);
        };
        const p1 = param(st);
        const p2 = param(st + sw);
        let d = p2 - p1;
        if (sw > 0 && d < 0) d += 2 * Math.PI;
        if (sw < 0 && d > 0) d -= 2 * Math.PI;
        if (Math.abs(sw) >= 360) d = Math.sign(sw) * 2 * Math.PI;
        const ecx = cx - wr * Math.cos(p1);
        const ecy = cy - hr * Math.sin(p1);
        cubics(ellipseCubics(ecx, ecy, wr, hr, 0, p1, d));
        break;
      }
    }
  }
  flush();
  let text: EvaluatedGeometry['text'];
  if (g.textAreas) {
    const ta = pathTokens(g.textAreas).map((x) => f.param(x));
    if (ta.length >= 4) {
      if (g.logicalText) {
        const k = 96 / 2540;
        text = [ta[0] * k, ta[1] * k, w - ta[2] * k, h - ta[3] * k].map((n) => Math.round(n * 100) / 100) as [number, number, number, number];
      } else {
        const sx = w / vw;
        const sy = h / vh;
        text = [(ta[0] - vx) * sx, (ta[1] - vy) * sy, (vx + vw - ta[2]) * sx, (vy + vh - ta[3]) * sy].map((n) => Math.round(n * 100) / 100) as [number, number, number, number];
      }
    }
  }
  return { paths, text };
}
