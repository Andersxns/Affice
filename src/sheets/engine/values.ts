import { parseInput } from '../format/numfmt';

export type ErrCode = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#SPILL!' | '#CALC!' | '#CIRC!';

/** An Excel error value. Serialises to { error: '#N/A' } like stored cells. */
export class FErr {
  constructor(public readonly error: ErrCode) {}
  toString() {
    return this.error;
  }
}

export const ERR = {
  NULL: new FErr('#NULL!'),
  DIV0: new FErr('#DIV/0!'),
  VALUE: new FErr('#VALUE!'),
  REF: new FErr('#REF!'),
  NAME: new FErr('#NAME?'),
  NUM: new FErr('#NUM!'),
  NA: new FErr('#N/A'),
  SPILL: new FErr('#SPILL!'),
  CALC: new FErr('#CALC!'),
  CIRC: new FErr('#CIRC!'),
};

export function errFromCode(code: string): FErr {
  const hit = Object.values(ERR).find((e) => e.error === code.toUpperCase());
  return hit ?? ERR.VALUE;
}

export type Scalar = number | string | boolean | null | FErr;

/** A lazy reference to a block of cells on a sheet (sheet = sheet id). */
export interface RefValue {
  kind: 'ref';
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface Matrix {
  kind: 'matrix';
  data: Scalar[][];
}

/** LAMBDA values (created by LAMBDA, called by name or MAP/REDUCE/etc). */
export interface LambdaValue {
  kind: 'lambda';
  params: string[];
  body: unknown;
  scope: Map<string, Value>;
}

export type Value = Scalar | RefValue | Matrix | LambdaValue;

export const isErr = (v: unknown): v is FErr => v instanceof FErr || (typeof v === 'object' && v !== null && 'error' in v && typeof (v as { error: unknown }).error === 'string' && !('kind' in v));
export const isRef = (v: unknown): v is RefValue => typeof v === 'object' && v !== null && (v as RefValue).kind === 'ref';
export const isMatrix = (v: unknown): v is Matrix => typeof v === 'object' && v !== null && (v as Matrix).kind === 'matrix';
export const isLambda = (v: unknown): v is LambdaValue => typeof v === 'object' && v !== null && (v as LambdaValue).kind === 'lambda';

export function asErr(v: unknown): FErr {
  if (v instanceof FErr) return v;
  return errFromCode((v as { error: string }).error);
}

export const matrix = (data: Scalar[][]): Matrix => ({ kind: 'matrix', data });

/* ------------------------------------------------------------- coercion */

export function toNumber(v: Scalar): number | FErr {
  if (typeof v === 'number') return Number.isFinite(v) ? v : ERR.NUM;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) return asErr(v);
  const s = v.trim();
  if (s === '') return ERR.VALUE;
  const p = parseInput(s);
  if (typeof p.value === 'number') return p.value;
  return ERR.VALUE;
}

/** Number → text the way Excel converts in formulas (up to 15 significant digits). */
export function numToText(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const abs = Math.abs(n);
  if (abs !== 0 && (abs >= 1e15 || abs < 1e-5)) {
    const [m, e] = n.toExponential(14).split('e');
    const mant = m.replace(/\.?0+$/, '');
    const exp = Number(e);
    return `${mant}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  return String(parseFloat(n.toPrecision(15)));
}

export function toText(v: Scalar): string | FErr {
  if (typeof v === 'string') return v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numToText(v);
  return asErr(v);
}

export function toBool(v: Scalar): boolean | FErr {
  if (typeof v === 'boolean') return v;
  if (v === null) return false;
  if (typeof v === 'number') return v !== 0;
  if (isErr(v)) return asErr(v);
  const s = v.trim().toUpperCase();
  if (s === 'TRUE') return true;
  if (s === 'FALSE') return false;
  return ERR.VALUE;
}

/** Excel's ordering for comparisons: numbers < text < booleans; text is case-insensitive. */
export function compare(a: Scalar, b: Scalar): number {
  const rank = (x: Scalar) => (typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : typeof x === 'boolean' ? 2 : 0);
  let x = a;
  let y = b;
  if (x === null) x = typeof y === 'string' ? '' : typeof y === 'boolean' ? false : 0;
  if (y === null) y = typeof x === 'string' ? '' : typeof x === 'boolean' ? false : 0;
  const ra = rank(x);
  const rb = rank(y);
  if (ra !== rb) return ra - rb;
  if (typeof x === 'number' && typeof y === 'number') return x === y ? 0 : x < y ? -1 : 1;
  if (typeof x === 'string' && typeof y === 'string') {
    const p = x.toLowerCase();
    const q = y.toLowerCase();
    return p === q ? 0 : p < q ? -1 : 1;
  }
  if (typeof x === 'boolean' && typeof y === 'boolean') return x === y ? 0 : x ? 1 : -1;
  return 0;
}

/* ------------------------------------------------------------- criteria */

function wildcardToRegex(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) {
      re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    } else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'i');
}

/** Builds a matcher for COUNTIF/SUMIF-style criteria ("&gt;5", "app*", "&lt;&gt;x", 10, TRUE…). */
export function makeCriteria(crit: Scalar): (v: Scalar) => boolean {
  if (typeof crit === 'number') return (v) => typeof v === 'number' ? v === crit : typeof v === 'string' && v.trim() !== '' && toNumber(v) === crit;
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit === null) return (v) => v === null || v === '';
  if (isErr(crit)) {
    const code = asErr(crit).error;
    return (v) => isErr(v) && asErr(v).error === code;
  }
  const s = String(crit);
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(s)!;
  const op = m[1] ?? '=';
  const rhs = m[2];
  const rhsNum = rhs.trim() !== '' ? toNumber(rhs) : ERR.VALUE;
  const numeric = typeof rhsNum === 'number';
  const rhsBool = rhs.toUpperCase() === 'TRUE' ? true : rhs.toUpperCase() === 'FALSE' ? false : null;
  if (op === '=' || op === '<>') {
    let test: (v: Scalar) => boolean;
    if (rhs === '') test = (v) => v === null || v === '';
    else if (numeric) test = (v) => (typeof v === 'number' ? v === rhsNum : typeof v === 'string' && toNumber(v) === rhsNum);
    else if (rhsBool !== null) test = (v) => v === rhsBool;
    else if (/[*?~]/.test(rhs)) {
      const re = wildcardToRegex(rhs);
      test = (v) => typeof v === 'string' && re.test(v);
    } else {
      const low = rhs.toLowerCase();
      test = (v) => typeof v === 'string' && v.toLowerCase() === low;
    }
    return op === '=' ? test : (v) => !test(v);
  }
  return (v) => {
    if (numeric) {
      if (typeof v !== 'number') return false;
      const c = v - (rhsNum as number);
      return op === '<' ? c < 0 : op === '<=' ? c <= 0 : op === '>' ? c > 0 : c >= 0;
    }
    if (typeof v !== 'string') return false;
    const c = compare(v, rhs);
    return op === '<' ? c < 0 : op === '<=' ? c <= 0 : op === '>' ? c > 0 : c >= 0;
  };
}
