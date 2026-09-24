import type { Node } from './parser';
import {
  asErr,
  compare,
  ERR,
  FErr,
  isErr,
  isLambda,
  isMatrix,
  isRef,
  matrix,
  toNumber,
  toText,
  type LambdaValue,
  type Matrix,
  type RefValue,
  type Scalar,
  type Value,
} from './values';
import { getFunction, type FnDef } from './functions/registry';

/** Everything a formula can see of the workbook while it is evaluated. */
export interface EvalContext {
  /** Sheet id and position of the cell being evaluated. */
  sheet: number;
  row: number;
  col: number;
  /** Computed value of a cell (evaluates stale formula cells on demand). */
  cell(sheet: number, r: number, c: number): Scalar;
  /** Visits the non-empty cells of a reference, in no particular order. */
  each(ref: RefValue, cb: (v: Scalar, r: number, c: number) => void): void;
  /** Index of the last used row / column of a sheet (-1 when empty). */
  extent(sheet: number): { rows: number; cols: number };
  sheetId(name: string): number | undefined;
  sheetName(id: number): string | undefined;
  sheetIds(): number[];
  definedName(name: string, sheet: number): Node | undefined;
  formulaText(sheet: number, r: number, c: number): string | undefined;
  /** Range covered by the spilled array anchored at a cell. */
  spillRange(sheet: number, r: number, c: number): RefValue | undefined;
  /** Number format of a cell (for CELL("format") / TEXT helpers). */
  cellFormat?(sheet: number, r: number, c: number): string | undefined;
  /** Whether a row is hidden by the user or by a filter (SUBTOTAL/AGGREGATE). */
  rowState?(sheet: number, r: number): 'hidden' | 'filtered' | undefined;
  now(): number;
  random(): number;
  markVolatile(): void;
}

export type Scope = Map<string, Value>;

const OMITTED_PREFIX = '\u0000omitted:';

/* ================================================================ refs */

/** Whole rows/columns are clamped to the used area before they are materialised. */
const CLAMP_ROWS = 65536;
const CLAMP_COLS = 1024;

export function clampRef(ref: RefValue, ctx: EvalContext): RefValue {
  let { r2, c2 } = ref;
  if (r2 - ref.r1 + 1 > CLAMP_ROWS || c2 - ref.c1 + 1 > CLAMP_COLS) {
    const ext = ctx.extent(ref.sheet);
    if (r2 - ref.r1 + 1 > CLAMP_ROWS) r2 = Math.max(ref.r1, Math.min(r2, ext.rows));
    if (c2 - ref.c1 + 1 > CLAMP_COLS) c2 = Math.max(ref.c1, Math.min(c2, ext.cols));
  }
  return r2 === ref.r2 && c2 === ref.c2 ? ref : { ...ref, r2, c2 };
}

export function refToMatrix(ref: RefValue, ctx: EvalContext): Matrix {
  const rg = clampRef(ref, ctx);
  const rows = rg.r2 - rg.r1 + 1;
  const cols = rg.c2 - rg.c1 + 1;
  const data: Scalar[][] = new Array(rows);
  for (let i = 0; i < rows; i++) {
    const row: Scalar[] = new Array(cols);
    for (let j = 0; j < cols; j++) row[j] = ctx.cell(rg.sheet, rg.r1 + i, rg.c1 + j);
    data[i] = row;
  }
  return matrix(data);
}

export const isSingleCell = (ref: RefValue) => ref.r1 === ref.r2 && ref.c1 === ref.c2;

/** Resolves references: single cells become scalars, larger ranges become matrices. */
export function deref(v: Value, ctx: EvalContext): Scalar | Matrix | LambdaValue {
  if (isRef(v)) return isSingleCell(v) ? ctx.cell(v.sheet, v.r1, v.c1) : refToMatrix(v, ctx);
  return v;
}

/** Any value as a 2D array (scalars become 1×1). */
export function toGrid(v: Value, ctx: EvalContext): Scalar[][] {
  if (isRef(v)) return refToMatrix(v, ctx).data;
  if (isMatrix(v)) return v.data;
  if (isLambda(v)) return [[ERR.CALC]];
  return [[v]];
}

/** Collapses a value to one scalar (top-left of arrays). */
export function toScalar(v: Value, ctx: EvalContext): Scalar {
  if (isRef(v)) return ctx.cell(v.sheet, v.r1, v.c1);
  if (isMatrix(v)) return v.data[0]?.[0] ?? null;
  if (isLambda(v)) return ERR.CALC;
  return v;
}

/** Excel's implicit intersection (the @ operator). */
export function implicitIntersect(v: Value, ctx: EvalContext): Scalar {
  if (isRef(v)) {
    if (isSingleCell(v)) return ctx.cell(v.sheet, v.r1, v.c1);
    if (v.c1 === v.c2 && ctx.row >= v.r1 && ctx.row <= v.r2) return ctx.cell(v.sheet, ctx.row, v.c1);
    if (v.r1 === v.r2 && ctx.col >= v.c1 && ctx.col <= v.c2) return ctx.cell(v.sheet, v.r1, ctx.col);
    return ERR.VALUE;
  }
  if (isMatrix(v)) return v.data[0]?.[0] ?? null;
  if (isLambda(v)) return ERR.CALC;
  return v;
}

/* =========================================================== broadcast */

export function matDims(m: Matrix): [number, number] {
  return [m.data.length, m.data[0]?.length ?? 0];
}

/** Element of an array under Excel's broadcasting rules (#N/A outside). */
export function at(m: Matrix, r: number, c: number): Scalar {
  const rows = m.data.length;
  const cols = m.data[0]?.length ?? 0;
  const i = rows === 1 ? 0 : r;
  const j = cols === 1 ? 0 : c;
  if (i >= rows || j >= cols) return ERR.NA;
  return m.data[i][j];
}

export function broadcastDims(items: (Scalar | Matrix)[]): [number, number] {
  let R = 1;
  let C = 1;
  for (const it of items) {
    if (!isMatrix(it)) continue;
    const [r, c] = matDims(it);
    if (r > R) R = r;
    if (c > C) C = c;
  }
  return [R, C];
}

export function mapMatrix(m: Matrix, fn: (v: Scalar) => Scalar): Matrix {
  return matrix(m.data.map((row) => row.map(fn)));
}

function lift1(v: Scalar | Matrix, fn: (x: Scalar) => Scalar): Scalar | Matrix {
  return isMatrix(v) ? mapMatrix(v, fn) : fn(v);
}

function lift2(a: Scalar | Matrix, b: Scalar | Matrix, fn: (x: Scalar, y: Scalar) => Scalar): Scalar | Matrix {
  if (!isMatrix(a) && !isMatrix(b)) return fn(a, b);
  const [R, C] = broadcastDims([a, b]);
  const out: Scalar[][] = new Array(R);
  for (let r = 0; r < R; r++) {
    const row: Scalar[] = new Array(C);
    for (let c = 0; c < C; c++) row[c] = fn(isMatrix(a) ? at(a, r, c) : a, isMatrix(b) ? at(b, r, c) : b);
    out[r] = row;
  }
  return matrix(out);
}

/* =========================================================== operators */

/** Excel snaps sums that are tiny relative to their operands to zero (0.1+0.2-0.3 → 0). */
function cleanSum(res: number, a: number, b: number): number {
  if (res === 0) return 0;
  const mag = Math.max(Math.abs(a), Math.abs(b));
  if (Math.abs(res) < mag * 1e-15) return 0;
  return res;
}

export function numEquals(a: number, b: number): boolean {
  if (a === b) return true;
  const mag = Math.max(Math.abs(a), Math.abs(b));
  return Math.abs(a - b) <= mag * 1e-15;
}

function arith(op: string, x: Scalar, y: Scalar): Scalar {
  if (isErr(x)) return asErr(x);
  if (isErr(y)) return asErr(y);
  const a = toNumber(x);
  if (typeof a !== 'number') return a;
  const b = toNumber(y);
  if (typeof b !== 'number') return b;
  let r: number;
  switch (op) {
    case '+':
      r = cleanSum(a + b, a, b);
      break;
    case '-':
      r = cleanSum(a - b, a, b);
      break;
    case '*':
      r = a * b;
      break;
    case '/':
      if (b === 0) return ERR.DIV0;
      r = a / b;
      break;
    case '^':
      if (a === 0 && b === 0) return ERR.NUM;
      if (a === 0 && b < 0) return ERR.DIV0;
      r = Math.pow(a, b);
      if (Number.isNaN(r) && a < 0) {
        // odd roots of negatives, e.g. (-8)^(1/3)
        const inv = 1 / b;
        if (Math.abs(inv - Math.round(inv)) < 1e-10 && Math.round(inv) % 2 !== 0) r = -Math.pow(-a, b);
      }
      break;
    default:
      return ERR.VALUE;
  }
  if (!Number.isFinite(r)) return ERR.NUM;
  return r;
}

function cmp(op: string, x: Scalar, y: Scalar): Scalar {
  if (isErr(x)) return asErr(x);
  if (isErr(y)) return asErr(y);
  let c: number;
  if (typeof x === 'number' && typeof y === 'number') c = numEquals(x, y) ? 0 : x < y ? -1 : 1;
  else c = compare(x, y);
  switch (op) {
    case '=':
      return c === 0;
    case '<>':
      return c !== 0;
    case '<':
      return c < 0;
    case '>':
      return c > 0;
    case '<=':
      return c <= 0;
    case '>=':
      return c >= 0;
  }
  return ERR.VALUE;
}

function concat(x: Scalar, y: Scalar): Scalar {
  const a = toText(x);
  if (typeof a !== 'string') return a;
  const b = toText(y);
  if (typeof b !== 'string') return b;
  if (a.length + b.length > 32767) return ERR.VALUE;
  return a + b;
}

function negate(x: Scalar): Scalar {
  if (isErr(x)) return asErr(x);
  const a = toNumber(x);
  return typeof a === 'number' ? (a === 0 ? 0 : -a) : a;
}

function percent(x: Scalar): Scalar {
  if (isErr(x)) return asErr(x);
  const a = toNumber(x);
  return typeof a === 'number' ? a / 100 : a;
}

/** Union of two references (the ':' range operator between expressions). */
function rangeOp(a: Value, b: Value): Value {
  if (!isRef(a) || !isRef(b)) return isErr(a) ? a : isErr(b) ? b : ERR.VALUE;
  if (a.sheet !== b.sheet) return ERR.VALUE;
  return { kind: 'ref', sheet: a.sheet, r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) };
}

/* ========================================================== evaluation */

let lambdaDepth = 0;
const MAX_LAMBDA_DEPTH = 200;

function resolveSheet(name: string | undefined, ctx: EvalContext): number | FErr {
  if (name === undefined) return ctx.sheet;
  const id = ctx.sheetId(name);
  return id === undefined ? ERR.REF : id;
}

export function evaluate(node: Node, ctx: EvalContext, scope?: Scope): Value {
  switch (node.t) {
    case 'num':
      return node.v;
    case 'str':
      return node.v;
    case 'bool':
      return node.v;
    case 'err':
      return errFromText(node.v);
    case 'missing':
      return null;
    case 'ref': {
      const sheet = resolveSheet(node.sheet, ctx);
      if (typeof sheet !== 'number') return sheet;
      return { kind: 'ref', sheet, r1: node.r, c1: node.c, r2: node.r, c2: node.c };
    }
    case 'range': {
      const sheet = resolveSheet(node.sheet, ctx);
      if (typeof sheet !== 'number') return sheet;
      return { kind: 'ref', sheet, r1: node.r1, c1: node.c1, r2: node.r2, c2: node.c2 };
    }
    case 'spill': {
      const base = evaluate(node.ref, ctx, scope);
      if (!isRef(base)) return isErr(base) ? base : ERR.REF;
      return ctx.spillRange(base.sheet, base.r1, base.c1) ?? ERR.REF;
    }
    case 'name':
      return evalName(node.name, node.sheet, ctx, scope);
    case 'func':
      return callFunction(node.name, node.args, ctx, scope);
    case 'call': {
      const callee = evaluate(node.callee, ctx, scope);
      if (!isLambda(callee)) return isErr(callee) ? callee : ERR.VALUE;
      const args = node.args.map((a) => evaluate(a, ctx, scope));
      return callLambda(callee, args, ctx);
    }
    case 'unary': {
      const v = evaluate(node.arg, ctx, scope);
      if (node.op === '@') return implicitIntersect(v, ctx);
      if (node.op === '+') return v;
      const d = deref(v, ctx);
      if (isLambda(d)) return ERR.CALC;
      return lift1(d, negate);
    }
    case 'postfix': {
      const d = deref(evaluate(node.arg, ctx, scope), ctx);
      if (isLambda(d)) return ERR.CALC;
      return lift1(d, percent);
    }
    case 'bin': {
      const l = evaluate(node.left, ctx, scope);
      const r = evaluate(node.right, ctx, scope);
      if (node.op === ':') return rangeOp(l, r);
      const a = deref(l, ctx);
      const b = deref(r, ctx);
      if (isLambda(a) || isLambda(b)) return ERR.CALC;
      switch (node.op) {
        case '+':
        case '-':
        case '*':
        case '/':
        case '^':
          return lift2(a, b, (x, y) => arith(node.op, x, y));
        case '&':
          return lift2(a, b, concat);
        default:
          return lift2(a, b, (x, y) => cmp(node.op, x, y));
      }
    }
    case 'array': {
      const width = Math.max(...node.rows.map((r) => r.length));
      const data = node.rows.map((row) => {
        const out: Scalar[] = row.map((cell) => {
          const v = evaluate(cell, ctx, scope);
          return isRef(v) || isMatrix(v) || isLambda(v) ? ERR.VALUE : v;
        });
        while (out.length < width) out.push(ERR.NA);
        return out;
      });
      return matrix(data);
    }
  }
}

function errFromText(code: string): FErr {
  for (const e of Object.values(ERR)) if (e.error === code) return e;
  return new FErr(code as FErr['error']);
}

function evalName(name: string, sheetName: string | undefined, ctx: EvalContext, scope?: Scope): Value {
  if (!sheetName && scope) {
    const hit = scope.get(name);
    if (hit !== undefined) return hit;
  }
  let sheet = ctx.sheet;
  if (sheetName) {
    const id = ctx.sheetId(sheetName);
    if (id === undefined) return ERR.REF;
    sheet = id;
  }
  const def = ctx.definedName(name, sheet);
  if (def) return evaluate(def, ctx);
  // eta-reduced built-ins, e.g. BYROW(A1:C3, SUM)
  const fn = getFunction(name);
  if (fn && !fn.lazy) return { kind: 'lambda', params: [], body: { builtin: fn.name }, scope: new Map() };
  return ERR.NAME;
}

/* ============================================================ functions */

export type ParamKind = 'v' | 'e' | 'r';

/**
 * Kind of the i-th parameter. `args` holds one letter per parameter and the last
 * letter repeats; a '|' starts a repeating group instead (SUMIFS = 'r|rv').
 */
export function paramKind(def: FnDef, i: number): ParamKind {
  const k = def.args;
  if (!k) return 'r';
  const bar = k.indexOf('|');
  if (bar < 0) return (i < k.length ? k[i] : k[k.length - 1]) as ParamKind;
  if (i < bar) return k[i] as ParamKind;
  const rep = k.length - bar - 1;
  return k[bar + 1 + ((i - bar) % rep)] as ParamKind;
}

export function callFunction(name: string, argNodes: Node[], ctx: EvalContext, scope?: Scope): Value {
  const def = getFunction(name);
  if (!def) {
    // LET-bound or defined-name LAMBDA called by name
    let target: Value | undefined = scope?.get(name);
    if (target === undefined) {
      const dn = ctx.definedName(name, ctx.sheet);
      if (dn) target = evaluate(dn, ctx);
    }
    if (target !== undefined && isLambda(target)) {
      const args = argNodes.map((a) => (a.t === 'missing' ? null : evaluate(a, ctx, scope)));
      return callLambda(target, args, ctx);
    }
    return ERR.NAME;
  }
  if (argNodes.length < def.min || argNodes.length > def.max) return ERR.VALUE;
  if (def.volatile) ctx.markVolatile();
  if (def.lazy) {
    try {
      return def.lazy(argNodes, ctx, scope);
    } catch (e) {
      return caught(e);
    }
  }
  const vals: Value[] = new Array(argNodes.length);
  for (let i = 0; i < argNodes.length; i++) {
    const a = argNodes[i];
    vals[i] = a.t === 'missing' ? null : evaluate(a, ctx, scope);
  }
  return invoke(def, vals, ctx);
}

function caught(e: unknown): Value {
  if (e instanceof FErr) return e;
  if (e instanceof RangeError) return ERR.NUM;
  if (isErr(e)) return asErr(e);
  console.warn('[sheets] function failed', e);
  return ERR.VALUE;
}

/** Calls a function definition with evaluated arguments, lifting scalar parameters over arrays. */
export function invoke(def: FnDef, vals: Value[], ctx: EvalContext): Value {
  const impl = def.fn!;
  const prepared: Value[] = new Array(vals.length);
  let lifted = false;
  for (let i = 0; i < vals.length; i++) {
    const kind = paramKind(def, i);
    let v = vals[i];
    if (kind !== 'r') {
      if (isRef(v)) v = isSingleCell(v) ? ctx.cell(v.sheet, v.r1, v.c1) : refToMatrix(v, ctx);
      if (isMatrix(v)) {
        if (v.data.length === 1 && v.data[0].length === 1) v = v.data[0][0];
        else lifted = true;
      } else if (kind === 'v' && isErr(v)) return asErr(v);
    }
    prepared[i] = v;
  }
  if (!lifted) {
    try {
      return impl(prepared, ctx);
    } catch (e) {
      return caught(e);
    }
  }
  const liftIdx: number[] = [];
  const mats: Matrix[] = [];
  prepared.forEach((v, i) => {
    if (paramKind(def, i) !== 'r' && isMatrix(v)) {
      liftIdx.push(i);
      mats.push(v);
    }
  });
  const [R, C] = broadcastDims(mats);
  const out: Scalar[][] = new Array(R);
  const args = prepared.slice();
  for (let r = 0; r < R; r++) {
    const row: Scalar[] = new Array(C);
    for (let c = 0; c < C; c++) {
      let err: FErr | null = null;
      for (let k = 0; k < liftIdx.length; k++) {
        const x = at(mats[k], r, c);
        if (paramKind(def, liftIdx[k]) === 'v' && isErr(x)) {
          err = asErr(x);
          break;
        }
        args[liftIdx[k]] = x;
      }
      if (err) {
        row[c] = err;
        continue;
      }
      let res: Value;
      try {
        res = impl(args, ctx);
      } catch (e) {
        res = caught(e);
      }
      row[c] = toScalar(res, ctx);
    }
    out[r] = row;
  }
  return matrix(out);
}

/* ============================================================== lambdas */

export function makeLambda(params: string[], body: Node, scope?: Scope): LambdaValue {
  return { kind: 'lambda', params, body, scope: scope ? new Map(scope) : new Map() };
}

export function callLambda(fn: LambdaValue, args: Value[], ctx: EvalContext): Value {
  const body = fn.body as Node | { builtin: string };
  if ('builtin' in body) {
    const def = getFunction(body.builtin);
    if (!def?.fn) return ERR.VALUE;
    if (args.length < def.min || args.length > def.max) return ERR.VALUE;
    return invoke(def, args, ctx);
  }
  if (args.length > fn.params.length) return ERR.VALUE;
  if (lambdaDepth >= MAX_LAMBDA_DEPTH) return ERR.NUM;
  const scope: Scope = new Map(fn.scope);
  fn.params.forEach((p, i) => {
    if (i < args.length) scope.set(p, args[i]);
    else {
      scope.set(p, null);
      scope.set(OMITTED_PREFIX + p, true);
    }
  });
  lambdaDepth++;
  try {
    return evaluate(body, ctx, scope);
  } catch (e) {
    return caught(e);
  } finally {
    lambdaDepth--;
  }
}

export function isOmitted(name: string, scope?: Scope): boolean {
  return !!scope?.get(OMITTED_PREFIX + name);
}

/* ============================================================== results */

/** Converts the value of a top-level formula into what a cell stores (scalar or spill array). */
export function finalize(v: Value, ctx: EvalContext): Scalar | Matrix {
  if (isRef(v)) {
    if (isSingleCell(v)) return normalizeScalar(ctx.cell(v.sheet, v.r1, v.c1), true);
    const m = refToMatrix(v, ctx);
    return matrix(m.data.map((row) => row.map((x) => normalizeScalar(x, true))));
  }
  if (isLambda(v)) return ERR.CALC;
  if (isMatrix(v)) {
    if (!v.data.length || !v.data[0]?.length) return ERR.CALC;
    if (v.data.length === 1 && v.data[0].length === 1) return normalizeScalar(v.data[0][0], false);
    return matrix(v.data.map((row) => row.map((x) => normalizeScalar(x, false))));
  }
  return normalizeScalar(v, false);
}

/** A formula that returns a blank cell shows 0, like Excel. */
function normalizeScalar(v: Scalar, fromRef: boolean): Scalar {
  if (v === null) return fromRef ? 0 : 0;
  if (typeof v === 'number') return Number.isFinite(v) ? (Object.is(v, -0) ? 0 : v) : ERR.NUM;
  return v;
}
