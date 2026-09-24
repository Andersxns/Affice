import { asErr, ERR, FErr, isErr, isLambda, isMatrix, isRef, makeCriteria, matrix, toBool, toNumber, toText, type Matrix, type Scalar, type Value } from '../values';
import { clampRef, refToMatrix, toGrid, type EvalContext } from '../evaluator';
import { define, type Category, type FnDef } from './registry';

/* ============================================================ coercion */

/** Scalar → number, throwing Excel errors (blank = 0). */
export function num(v: Scalar | undefined): number {
  if (v === undefined || v === null) return 0;
  if (typeof v === 'number') return v;
  const n = toNumber(v);
  if (typeof n !== 'number') throw n;
  return n;
}

/** Truncated integer argument. */
export function int(v: Scalar | undefined): number {
  return Math.trunc(num(v));
}

export function str(v: Scalar | undefined): string {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string') return v;
  const t = toText(v);
  if (typeof t !== 'string') throw t;
  return t;
}

export function bool(v: Scalar | undefined): boolean {
  if (v === undefined || v === null) return false;
  const b = toBool(v);
  if (typeof b !== 'boolean') throw b;
  return b;
}

/** Optional numeric parameter: omitted → default, empty (,,) → 0. */
export function optNum(args: Scalar[], i: number, def: number): number {
  if (i >= args.length) return def;
  return num(args[i]);
}

export function optInt(args: Scalar[], i: number, def: number): number {
  if (i >= args.length) return def;
  return int(args[i]);
}

export function optBool(args: Scalar[], i: number, def: boolean): boolean {
  if (i >= args.length) return def;
  return bool(args[i]);
}

export function optStr(args: Scalar[], i: number, def: string): string {
  if (i >= args.length) return def;
  return str(args[i]);
}

/** Omitted or empty (,,) → default. */
export function optNumBlank(args: Scalar[], i: number, def: number): number {
  if (i >= args.length || args[i] === null) return def;
  return num(args[i]);
}

export function check(cond: boolean, err: FErr = ERR.NUM): void {
  if (!cond) throw err;
}

export function finite(n: number): number {
  if (!Number.isFinite(n)) throw ERR.NUM;
  return n;
}

/* ============================================================ iteration */

/**
 * Visits every value of an argument. References visit their non-empty cells,
 * arrays visit every element; `direct` is true for scalars typed as arguments.
 */
export function eachValue(a: Value, ctx: EvalContext, cb: (v: Scalar, direct: boolean) => void): void {
  if (isRef(a)) ctx.each(a, (v) => cb(v, false));
  else if (isMatrix(a)) {
    for (const row of a.data) for (const v of row) cb(v, false);
  } else if (isLambda(a)) throw ERR.CALC;
  else cb(a, true);
}

export function eachArg(args: Value[], ctx: EvalContext, cb: (v: Scalar, direct: boolean) => void): void {
  for (const a of args) eachValue(a, ctx, cb);
}

/**
 * Numbers from arguments using Excel's aggregate rules.
 * mode 'n': references/arrays contribute numbers only (text/bools ignored);
 * mode 'a': references contribute text as 0 and booleans as 0/1 (the …A functions).
 * Typed arguments are always coerced; errors propagate.
 */
export function collectNums(args: Value[], ctx: EvalContext, mode: 'n' | 'a' = 'n'): number[] {
  const out: number[] = [];
  eachArg(args, ctx, (v, direct) => {
    if (typeof v === 'number') out.push(v);
    else if (v === null) {
      if (direct) out.push(0);
    } else if (isErr(v)) throw asErr(v);
    else if (direct) out.push(num(v));
    else if (mode === 'a') out.push(typeof v === 'boolean' ? (v ? 1 : 0) : 0);
  });
  return out;
}

/** Row-major 2D grid of any value (references are clamped to the used area). */
export function grid(v: Value, ctx: EvalContext): Scalar[][] {
  return toGrid(v, ctx);
}

export function flat(v: Value, ctx: EvalContext): Scalar[] {
  const g = toGrid(v, ctx);
  if (g.length === 1) return g[0].slice();
  const out: Scalar[] = [];
  for (const row of g) for (const x of row) out.push(x);
  return out;
}

export function dimsOf(v: Value, ctx: EvalContext): [number, number] {
  if (isRef(v)) {
    const rg = clampRef(v, ctx);
    return [rg.r2 - rg.r1 + 1, rg.c2 - rg.c1 + 1];
  }
  if (isMatrix(v)) return [v.data.length, v.data[0]?.length ?? 0];
  return [1, 1];
}

/** Full (unclamped) size of a value. */
export function fullDims(v: Value): [number, number] {
  if (isRef(v)) return [v.r2 - v.r1 + 1, v.c2 - v.c1 + 1];
  if (isMatrix(v)) return [v.data.length, v.data[0]?.length ?? 0];
  return [1, 1];
}

export function asMatrix(v: Value, ctx: EvalContext): Matrix {
  if (isRef(v)) return refToMatrix(v, ctx);
  if (isMatrix(v)) return v;
  return matrix([[isLambda(v) ? ERR.CALC : v]]);
}

/** Numbers of an array argument keeping positions (non-numbers → null). */
export function numGrid(v: Value, ctx: EvalContext): (number | null)[][] {
  return grid(v, ctx).map((row) =>
    row.map((x) => {
      if (isErr(x)) throw asErr(x);
      return typeof x === 'number' ? x : null;
    }),
  );
}

/** Paired numeric data (x/y) ignoring pairs where either side isn't a number. */
export function pairedNums(a: Value, b: Value, ctx: EvalContext): [number[], number[]] {
  const xa = flat(a, ctx);
  const xb = flat(b, ctx);
  if (xa.length !== xb.length) throw ERR.NA;
  const p: number[] = [];
  const q: number[] = [];
  for (let i = 0; i < xa.length; i++) {
    const x = xa[i];
    const y = xb[i];
    if (isErr(x)) throw asErr(x);
    if (isErr(y)) throw asErr(y);
    if (typeof x === 'number' && typeof y === 'number') {
      p.push(x);
      q.push(y);
    }
  }
  return [p, q];
}

/* ============================================================= criteria */

/**
 * Iterates positions of the first range where every (range, criterion) pair matches.
 * All ranges must share the same size (SUMIFS rules).
 */
export function matchIfs(pairs: [Value, Scalar][], ctx: EvalContext, cb: (r: number, c: number) => void): [number, number] {
  const grids = pairs.map(([rg]) => grid(rg, ctx));
  const tests = pairs.map(([, crit]) => makeCriteria(crit));
  const R = grids[0].length;
  const C = grids[0][0]?.length ?? 0;
  for (const g of grids) if (g.length !== R || (g[0]?.length ?? 0) !== C) throw ERR.VALUE;
  for (let r = 0; r < R; r++) {
    for (let c = 0; c < C; c++) {
      let ok = true;
      for (let k = 0; k < grids.length; k++) {
        if (!tests[k](grids[k][r][c])) {
          ok = false;
          break;
        }
      }
      if (ok) cb(r, c);
    }
  }
  return [R, C];
}

/** A range resized to the given size from its top-left corner (SUMIF's sum_range rule). */
export function resizedGrid(v: Value, rows: number, cols: number, ctx: EvalContext): Scalar[][] {
  if (isRef(v)) {
    const ref = { ...v, r2: v.r1 + rows - 1, c2: v.c1 + cols - 1 };
    return refToMatrix(ref, ctx).data;
  }
  const g = grid(v, ctx);
  const out: Scalar[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) row.push(g[r]?.[c] ?? null);
    out.push(row);
  }
  return out;
}

/* ============================================================= numerics */

/** Rounds half away from zero at 15 significant digits, like Excel. */
export function roundTo(n: number, digits: number): number {
  if (!Number.isFinite(n)) return n;
  const d = Math.trunc(digits);
  const sign = n < 0 ? -1 : 1;
  const a = Math.abs(n);
  if (d >= 0) {
    const f = Math.pow(10, d);
    const x = parseFloat((a * f).toPrecision(15));
    return (sign * Math.round(x)) / f;
  }
  const f = Math.pow(10, -d);
  const x = parseFloat((a / f).toPrecision(15));
  return sign * Math.round(x) * f;
}

export function roundUp(n: number, digits: number): number {
  const d = Math.trunc(digits);
  const sign = n < 0 ? -1 : 1;
  const a = Math.abs(n);
  const f = Math.pow(10, d);
  const x = parseFloat((a * f).toPrecision(15));
  return (sign * Math.ceil(x)) / f;
}

export function roundDown(n: number, digits: number): number {
  const d = Math.trunc(digits);
  const sign = n < 0 ? -1 : 1;
  const a = Math.abs(n);
  const f = Math.pow(10, d);
  const x = parseFloat((a * f).toPrecision(15));
  return (sign * Math.floor(x)) / f;
}

/** Removes binary noise (e.g. 0.30000000000000004 → 0.3) before integer tests. */
export function clean(n: number): number {
  return parseFloat(n.toPrecision(15));
}

export function sum(xs: number[]): number {
  // Neumaier summation keeps long columns of decimals exact
  let s = 0;
  let c = 0;
  for (const x of xs) {
    const t = s + x;
    if (Math.abs(s) >= Math.abs(x)) c += s - t + x;
    else c += x - t + s;
    s = t;
  }
  return s + c;
}

export function mean(xs: number[]): number {
  if (!xs.length) throw ERR.DIV0;
  return sum(xs) / xs.length;
}

export function variance(xs: number[], sample: boolean): number {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) throw ERR.DIV0;
  const m = mean(xs);
  let ss = 0;
  for (const x of xs) ss += (x - m) * (x - m);
  return ss / (sample ? n - 1 : n);
}

export function sortedNums(xs: number[]): number[] {
  return xs.slice().sort((a, b) => a - b);
}

/* ============================================================ builders */

type Impl = NonNullable<FnDef['fn']>;

/** Shorthand for a function with scalar parameters. */
export function fn(name: string, category: Category, min: number, max: number, args: string, syntax: string, desc: string, impl: Impl, extra: Partial<FnDef> = {}): void {
  define({ name, category, min, max, args, syntax, desc, fn: impl, ...extra });
}

/** A one-argument numeric function. */
export function math1(name: string, desc: string, f: (x: number) => number, category: Category = 'Math'): void {
  fn(name, category, 1, 1, 'v', 'number', desc, ([x]) => finite(f(num(x))));
}

export { matrix, ERR, isErr, asErr, isRef, isMatrix, isLambda };
