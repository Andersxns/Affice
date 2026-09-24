import * as formulajs from '@formulajs/formulajs';
import { besseli, besselj, besselk, bessely } from 'bessel';
import jStat from 'jstat';
import { ERR, errFromCode, isErr, asErr, type Scalar, type Value } from '../values';
import type { EvalContext } from '../evaluator';
import { check, fn, int, num, str } from './helpers';

const E = 'Engineering' as const;

/* ------------------------------------------------------ base conversion */

const LIMITS = { 2: 10, 8: 10, 16: 10 } as const;

function fromBase(text: Scalar, base: 2 | 8 | 16): number {
  const s = (typeof text === 'number' ? String(text) : str(text)).trim().toUpperCase();
  if (s === '') return 0;
  const re = base === 2 ? /^[01]{1,10}$/ : base === 8 ? /^[0-7]{1,10}$/ : /^[0-9A-F]{1,10}$/;
  if (!re.test(s)) throw ERR.NUM;
  let v = parseInt(s, base);
  const bits = base === 2 ? 10 : base === 8 ? 30 : 40;
  if (s.length === 10 && v >= 2 ** (bits - 1)) v -= 2 ** bits;
  return v;
}

function toBase(n: number, base: 2 | 8 | 16, places: number | undefined): string {
  const bits = base === 2 ? 10 : base === 8 ? 30 : 40;
  const v = Math.trunc(n);
  check(v >= -(2 ** (bits - 1)) && v < 2 ** (bits - 1));
  if (v < 0) return (2 ** bits + v).toString(base).toUpperCase();
  const s = v.toString(base).toUpperCase();
  if (places !== undefined) {
    check(places >= s.length && places <= LIMITS[base]);
    return s.padStart(places, '0');
  }
  return s;
}

const BASES: [string, 2 | 8 | 10 | 16][] = [
  ['BIN', 2],
  ['OCT', 8],
  ['DEC', 10],
  ['HEX', 16],
];
const LONG: Record<number, string> = { 2: 'binary', 8: 'octal', 10: 'decimal', 16: 'hexadecimal' };

for (const [fromName, fromBaseN] of BASES) {
  for (const [toName, toBaseN] of BASES) {
    if (fromName === toName) continue;
    const name = `${fromName}2${toName}`;
    const hasPlaces = toBaseN !== 10;
    fn(name, E, 1, hasPlaces ? 2 : 1, 'v', hasPlaces ? 'number, [places]' : 'number', `Converts a ${LONG[fromBaseN]} number to ${LONG[toBaseN]}.`, (args) => {
      const n = fromBaseN === 10 ? Math.trunc(num(args[0])) : fromBase(args[0], fromBaseN);
      if (toBaseN === 10) return n;
      const places = args.length > 1 && args[1] !== null ? int(args[1]) : undefined;
      return toBase(n, toBaseN, places);
    });
  }
}

/* ------------------------------------------------------------- bitwise */

function bitArg(v: Scalar): bigint {
  const n = num(v);
  check(n >= 0 && n < 2 ** 48 && Number.isInteger(n));
  return BigInt(n);
}

fn('BITAND', E, 2, 2, 'v', 'number1, number2', 'Returns a bitwise AND of two numbers.', ([a, b]) => Number(bitArg(a) & bitArg(b)));
fn('BITOR', E, 2, 2, 'v', 'number1, number2', 'Returns a bitwise OR of two numbers.', ([a, b]) => Number(bitArg(a) | bitArg(b)));
fn('BITXOR', E, 2, 2, 'v', 'number1, number2', 'Returns a bitwise exclusive OR of two numbers.', ([a, b]) => Number(bitArg(a) ^ bitArg(b)));
fn('BITLSHIFT', E, 2, 2, 'v', 'number, shift_amount', 'Returns a number shifted left by shift_amount bits.', ([a, s]) => {
  const sh = int(s);
  check(Math.abs(sh) <= 53);
  const r = sh >= 0 ? bitArg(a) << BigInt(sh) : bitArg(a) >> BigInt(-sh);
  check(r < 2n ** 48n);
  return Number(r);
});
fn('BITRSHIFT', E, 2, 2, 'v', 'number, shift_amount', 'Returns a number shifted right by shift_amount bits.', ([a, s]) => {
  const sh = int(s);
  check(Math.abs(sh) <= 53);
  const r = sh >= 0 ? bitArg(a) >> BigInt(sh) : bitArg(a) << BigInt(-sh);
  check(r < 2n ** 48n);
  return Number(r);
});

fn('DELTA', E, 1, 2, 'v', 'number1, [number2]', 'Tests whether two numbers are equal (returns 1 or 0).', (args) => (num(args[0]) === (args.length > 1 ? num(args[1]) : 0) ? 1 : 0));
fn('GESTEP', E, 1, 2, 'v', 'number, [step]', 'Tests whether a number is greater than or equal to a threshold value (returns 1 or 0).', (args) => (num(args[0]) >= (args.length > 1 ? num(args[1]) : 0) ? 1 : 0));

/* -------------------------------------------------------- error function */

fn('ERF', E, 1, 2, 'v', 'lower_limit, [upper_limit]', 'Returns the error function.', (args) => {
  const a = num(args[0]);
  if (args.length > 1 && args[1] !== null) return jStat.erf(num(args[1])) - jStat.erf(a);
  return jStat.erf(a);
});
fn('ERF.PRECISE', E, 1, 1, 'v', 'x', 'Returns the error function.', ([x]) => jStat.erf(num(x)));
fn('ERFC', E, 1, 1, 'v', 'x', 'Returns the complementary error function.', ([x]) => jStat.erfc(num(x)));
fn('ERFC.PRECISE', E, 1, 1, 'v', 'x', 'Returns the complementary error function.', ([x]) => jStat.erfc(num(x)));

/* --------------------------------------------------------------- bessel */

function besselFn(name: string, f: (x: number, n: number) => number, desc: string): void {
  fn(name, E, 2, 2, 'v', 'x, n', desc, ([x, n]) => {
    const order = Math.trunc(num(n));
    check(order >= 0);
    const r = f(num(x), order);
    if (!Number.isFinite(r)) throw ERR.NUM;
    return r;
  });
}
besselFn('BESSELI', besseli, 'Returns the modified Bessel function In(x).');
besselFn('BESSELJ', besselj, 'Returns the Bessel function Jn(x).');
besselFn('BESSELK', besselk, 'Returns the modified Bessel function Kn(x).');
besselFn('BESSELY', bessely, 'Returns the Bessel function Yn(x).');

/* ------------------------------------------ complex numbers & CONVERT */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const FJS = formulajs as unknown as Record<string, (...a: any[]) => any>;

function fromFjs(r: unknown): Scalar {
  if (r instanceof Error) return errFromCode(r.message);
  if (typeof r === 'number') return Number.isFinite(r) ? r : ERR.NUM;
  if (typeof r === 'string' || typeof r === 'boolean') return r;
  if (r === undefined || r === null) return 0;
  return ERR.VALUE;
}

function toFjs(v: Scalar): unknown {
  if (isErr(v)) throw asErr(v);
  return v === null ? undefined : v;
}

function wrapScalar(name: string, min: number, max: number, syntax: string, desc: string, target = name): void {
  const impl = FJS[target];
  if (typeof impl !== 'function') return;
  fn(name, E, min, max, 'v', syntax, desc, (args) => fromFjs(impl(...args.map(toFjs))));
}

wrapScalar('CONVERT', 3, 3, 'number, from_unit, to_unit', 'Converts a number from one measurement system to another (e.g. "mi" to "km", "C" to "F").');
wrapScalar('COMPLEX', 2, 3, 'real_num, i_num, [suffix]', 'Converts real and imaginary coefficients into a complex number.');
const IM1: [string, string][] = [
  ['IMABS', 'Returns the absolute value (modulus) of a complex number.'],
  ['IMAGINARY', 'Returns the imaginary coefficient of a complex number.'],
  ['IMARGUMENT', 'Returns the argument θ, an angle expressed in radians.'],
  ['IMCONJUGATE', 'Returns the complex conjugate of a complex number.'],
  ['IMCOS', 'Returns the cosine of a complex number.'],
  ['IMCOSH', 'Returns the hyperbolic cosine of a complex number.'],
  ['IMCOT', 'Returns the cotangent of a complex number.'],
  ['IMCSC', 'Returns the cosecant of a complex number.'],
  ['IMCSCH', 'Returns the hyperbolic cosecant of a complex number.'],
  ['IMEXP', 'Returns the exponential of a complex number.'],
  ['IMLN', 'Returns the natural logarithm of a complex number.'],
  ['IMLOG10', 'Returns the base-10 logarithm of a complex number.'],
  ['IMLOG2', 'Returns the base-2 logarithm of a complex number.'],
  ['IMREAL', 'Returns the real coefficient of a complex number.'],
  ['IMSEC', 'Returns the secant of a complex number.'],
  ['IMSECH', 'Returns the hyperbolic secant of a complex number.'],
  ['IMSIN', 'Returns the sine of a complex number.'],
  ['IMSINH', 'Returns the hyperbolic sine of a complex number.'],
  ['IMSQRT', 'Returns the square root of a complex number.'],
  ['IMTAN', 'Returns the tangent of a complex number.'],
];
for (const [name, desc] of IM1) wrapScalar(name, 1, 1, 'inumber', desc);
wrapScalar('IMDIV', 2, 2, 'inumber1, inumber2', 'Returns the quotient of two complex numbers.');
wrapScalar('IMPOWER', 2, 2, 'inumber, number', 'Returns a complex number raised to an integer power.');
wrapScalar('IMSUB', 2, 2, 'inumber1, inumber2', 'Returns the difference between two complex numbers.');

function wrapVariadic(name: string, desc: string): void {
  const impl = FJS[name];
  if (typeof impl !== 'function') return;
  fn(name, E, 1, 255, 'r', 'inumber1, [inumber2], …', desc, (args: Value[], ctx: EvalContext) => {
    const flat: unknown[] = [];
    for (const a of args) {
      const vals: Scalar[] = [];
      ctxEach(a, ctx, (v) => vals.push(v));
      for (const v of vals) flat.push(toFjs(v));
    }
    return fromFjs(impl(...flat));
  });
}

function ctxEach(a: Value, ctx: EvalContext, cb: (v: Scalar) => void): void {
  if (a && typeof a === 'object' && 'kind' in a) {
    if (a.kind === 'ref') ctx.each(a, (v) => cb(v));
    else if (a.kind === 'matrix') a.data.forEach((row) => row.forEach(cb));
    else throw ERR.VALUE;
  } else cb(a as Scalar);
}

wrapVariadic('IMSUM', 'Returns the sum of complex numbers.');
wrapVariadic('IMPRODUCT', 'Returns the product of complex numbers.');
