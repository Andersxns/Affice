import { ERR, isErr, asErr, makeCriteria, matrix, type Scalar, type Value } from '../values';
import type { EvalContext } from '../evaluator';
import {
  bool,
  check,
  clean,
  collectNums,
  eachArg,
  finite,
  flat,
  fn,
  grid,
  int,
  math1,
  matchIfs,
  num,
  optInt,
  optNum,
  resizedGrid,
  roundDown,
  roundTo,
  roundUp,
  str,
  sum,
} from './helpers';
import { aggregateBy } from './stats';

const M = 'Math' as const;

/* ------------------------------------------------------------ basics */

fn('SUM', M, 1, 255, 'r', 'number1, [number2], …', 'Adds all the numbers in a range of cells.', (args, ctx) => sum(collectNums(args, ctx)));

fn('PRODUCT', M, 1, 255, 'r', 'number1, [number2], …', 'Multiplies all the numbers given as arguments.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  if (!xs.length) return 0;
  return finite(xs.reduce((a, b) => a * b, 1));
});

fn('SUMSQ', M, 1, 255, 'r', 'number1, [number2], …', 'Returns the sum of the squares of the arguments.', (args, ctx) => finite(sum(collectNums(args, ctx).map((x) => x * x))));

fn('SUMPRODUCT', M, 1, 255, 'r', 'array1, [array2], …', 'Multiplies corresponding components in the given arrays and returns the sum of those products.', (args: Value[], ctx) => {
  const gs = args.map((a) => grid(a, ctx));
  const R = gs[0].length;
  const C = gs[0][0]?.length ?? 0;
  for (const g of gs) if (g.length !== R || (g[0]?.length ?? 0) !== C) throw ERR.VALUE;
  let s = 0;
  for (let r = 0; r < R; r++)
    for (let c = 0; c < C; c++) {
      let p = 1;
      for (const g of gs) {
        const v = g[r][c];
        if (isErr(v)) throw asErr(v);
        p *= typeof v === 'number' ? v : 0;
      }
      s += p;
    }
  return finite(s);
});

math1('ABS', 'Returns the absolute value of a number.', Math.abs);
math1('SIGN', 'Returns the sign of a number: 1, 0 or -1.', (x) => (x > 0 ? 1 : x < 0 ? -1 : 0));
math1('INT', 'Rounds a number down to the nearest integer.', (x) => Math.floor(clean(x)));
math1('SQRT', 'Returns a positive square root.', (x) => {
  check(x >= 0);
  return Math.sqrt(x);
});
math1('SQRTPI', 'Returns the square root of (number × pi).', (x) => {
  check(x >= 0);
  return Math.sqrt(x * Math.PI);
});
math1('EXP', 'Returns e raised to the power of a given number.', Math.exp);
math1('LN', 'Returns the natural logarithm of a number.', (x) => {
  check(x > 0);
  return Math.log(x);
});
math1('LOG10', 'Returns the base-10 logarithm of a number.', (x) => {
  check(x > 0);
  return Math.log10(x);
});
math1('DEGREES', 'Converts radians to degrees.', (x) => (x * 180) / Math.PI);
math1('RADIANS', 'Converts degrees to radians.', (x) => (x * Math.PI) / 180);
math1('EVEN', 'Rounds a number up (away from zero) to the nearest even integer.', (x) => {
  const s = x < 0 ? -1 : 1;
  let n = Math.ceil(Math.abs(clean(x)));
  if (n % 2) n += 1;
  return s * n;
});
math1('ODD', 'Rounds a number up (away from zero) to the nearest odd integer.', (x) => {
  const s = x < 0 ? -1 : 1;
  let n = Math.ceil(Math.abs(clean(x)));
  if (n % 2 === 0) n += 1;
  return s * n;
});
math1('FACT', 'Returns the factorial of a number.', (x) => {
  check(x >= 0);
  const n = Math.floor(x);
  check(n <= 170);
  let r = 1;
  for (let i = 2; i <= n; i++) r *= i;
  return r;
});
math1('FACTDOUBLE', 'Returns the double factorial of a number.', (x) => {
  check(x >= -1);
  const n = Math.floor(x);
  let r = 1;
  for (let i = n; i > 1; i -= 2) r *= i;
  return r;
});
math1('GAMMA', 'Returns the Gamma function value.', (x) => {
  check(!(x <= 0 && Number.isInteger(x)));
  return gamma(x);
}, 'Statistical');
math1('GAMMALN', 'Returns the natural logarithm of the gamma function.', (x) => {
  check(x > 0);
  return gammaln(x);
}, 'Statistical');
math1('GAMMALN.PRECISE', 'Returns the natural logarithm of the gamma function.', (x) => {
  check(x > 0);
  return gammaln(x);
}, 'Statistical');

fn('PI', M, 0, 0, 'v', '', 'Returns the value of pi.', () => Math.PI);

fn('POWER', M, 2, 2, 'v', 'number, power', 'Returns the result of a number raised to a power.', ([a, b]) => {
  const x = num(a);
  const y = num(b);
  if (x === 0 && y === 0) throw ERR.NUM;
  if (x === 0 && y < 0) throw ERR.DIV0;
  let r = Math.pow(x, y);
  if (Number.isNaN(r) && x < 0) {
    const inv = 1 / y;
    if (Math.abs(inv - Math.round(inv)) < 1e-10 && Math.round(inv) % 2 !== 0) r = -Math.pow(-x, y);
  }
  return finite(r);
});

fn('LOG', M, 1, 2, 'v', 'number, [base]', 'Returns the logarithm of a number to the base you specify (10 by default).', (args) => {
  const x = num(args[0]);
  const b = optNum(args, 1, 10);
  check(x > 0 && b > 0);
  if (b === 1) throw ERR.DIV0;
  return Math.log(x) / Math.log(b);
});

fn('MOD', M, 2, 2, 'v', 'number, divisor', 'Returns the remainder after a number is divided by a divisor.', ([a, b]) => {
  const n = num(a);
  const d = num(b);
  if (d === 0) throw ERR.DIV0;
  const r = n - d * Math.floor(n / d);
  return Math.abs(r - d) < Math.abs(d) * 1e-15 ? 0 : clean(r);
});

fn('QUOTIENT', M, 2, 2, 'v', 'numerator, denominator', 'Returns the integer portion of a division.', ([a, b]) => {
  const d = num(b);
  if (d === 0) throw ERR.DIV0;
  return Math.trunc(clean(num(a) / d));
});

/* ---------------------------------------------------------- rounding */

fn('ROUND', M, 2, 2, 'v', 'number, num_digits', 'Rounds a number to a specified number of digits.', ([a, b]) => roundTo(num(a), int(b)));
fn('ROUNDUP', M, 2, 2, 'v', 'number, num_digits', 'Rounds a number up, away from zero.', ([a, b]) => roundUp(num(a), int(b)));
fn('ROUNDDOWN', M, 2, 2, 'v', 'number, num_digits', 'Rounds a number down, toward zero.', ([a, b]) => roundDown(num(a), int(b)));
fn('TRUNC', M, 1, 2, 'v', 'number, [num_digits]', 'Truncates a number to an integer or to a number of decimals.', (args) => roundDown(num(args[0]), optInt(args, 1, 0)));

fn('MROUND', M, 2, 2, 'v', 'number, multiple', 'Returns a number rounded to the desired multiple.', ([a, b]) => {
  const n = num(a);
  const m = num(b);
  if (m === 0) return 0;
  if (n * m < 0) throw ERR.NUM;
  return clean(roundTo(n / m, 0) * m);
});

function ceilTo(x: number, sig: number): number {
  return clean(Math.ceil(clean(x / sig)) * sig);
}
function floorTo(x: number, sig: number): number {
  return clean(Math.floor(clean(x / sig)) * sig);
}

fn('CEILING', M, 1, 2, 'v', 'number, [significance]', 'Rounds a number up to the nearest multiple of significance.', (args) => {
  const x = num(args[0]);
  const s = optNum(args, 1, 1);
  if (s === 0) return 0;
  if (x > 0 && s < 0) throw ERR.NUM;
  if (x < 0 && s < 0) return -ceilTo(-x, -s);
  return ceilTo(x, s);
});
fn('CEILING.MATH', M, 1, 3, 'v', 'number, [significance], [mode]', 'Rounds a number up to the nearest integer or multiple of significance.', (args) => {
  const x = num(args[0]);
  const s = Math.abs(optNum(args, 1, 1)) || 0;
  const mode = optNum(args, 2, 0);
  if (s === 0) return 0;
  if (x < 0 && mode !== 0) return -ceilTo(-x, s);
  return ceilTo(x, s);
});
fn('CEILING.PRECISE', M, 1, 2, 'v', 'number, [significance]', 'Rounds a number up to the nearest multiple of significance, regardless of sign.', (args) => {
  const s = Math.abs(optNum(args, 1, 1));
  return s === 0 ? 0 : ceilTo(num(args[0]), s);
});
fn('ISO.CEILING', M, 1, 2, 'v', 'number, [significance]', 'Rounds a number up to the nearest multiple of significance.', (args) => {
  const s = Math.abs(optNum(args, 1, 1));
  return s === 0 ? 0 : ceilTo(num(args[0]), s);
});
fn('FLOOR', M, 1, 2, 'v', 'number, [significance]', 'Rounds a number down toward zero to the nearest multiple of significance.', (args) => {
  const x = num(args[0]);
  const s = optNum(args, 1, 1);
  if (s === 0) {
    if (x === 0) return 0;
    throw ERR.DIV0;
  }
  if (x > 0 && s < 0) throw ERR.NUM;
  if (x < 0 && s < 0) return -floorTo(-x, -s);
  return floorTo(x, s);
});
fn('FLOOR.MATH', M, 1, 3, 'v', 'number, [significance], [mode]', 'Rounds a number down to the nearest integer or multiple of significance.', (args) => {
  const x = num(args[0]);
  const s = Math.abs(optNum(args, 1, 1));
  const mode = optNum(args, 2, 0);
  if (s === 0) return 0;
  if (x < 0 && mode !== 0) return -floorTo(-x, s);
  return floorTo(x, s);
});
fn('FLOOR.PRECISE', M, 1, 2, 'v', 'number, [significance]', 'Rounds a number down to the nearest multiple of significance, regardless of sign.', (args) => {
  const s = Math.abs(optNum(args, 1, 1));
  return s === 0 ? 0 : floorTo(num(args[0]), s);
});

/* --------------------------------------------------------------- trig */

math1('SIN', 'Returns the sine of an angle (radians).', Math.sin);
math1('COS', 'Returns the cosine of an angle (radians).', Math.cos);
math1('TAN', 'Returns the tangent of an angle (radians).', Math.tan);
math1('ASIN', 'Returns the arcsine of a number.', (x) => {
  check(x >= -1 && x <= 1);
  return Math.asin(x);
});
math1('ACOS', 'Returns the arccosine of a number.', (x) => {
  check(x >= -1 && x <= 1);
  return Math.acos(x);
});
math1('ATAN', 'Returns the arctangent of a number.', Math.atan);
math1('SINH', 'Returns the hyperbolic sine of a number.', Math.sinh);
math1('COSH', 'Returns the hyperbolic cosine of a number.', Math.cosh);
math1('TANH', 'Returns the hyperbolic tangent of a number.', Math.tanh);
math1('ASINH', 'Returns the inverse hyperbolic sine of a number.', Math.asinh);
math1('ACOSH', 'Returns the inverse hyperbolic cosine of a number.', (x) => {
  check(x >= 1);
  return Math.acosh(x);
});
math1('ATANH', 'Returns the inverse hyperbolic tangent of a number.', (x) => {
  check(x > -1 && x < 1);
  return Math.atanh(x);
});
math1('COT', 'Returns the cotangent of an angle.', (x) => {
  if (x === 0) throw ERR.DIV0;
  return 1 / Math.tan(x);
});
math1('COTH', 'Returns the hyperbolic cotangent of a number.', (x) => {
  if (x === 0) throw ERR.DIV0;
  return 1 / Math.tanh(x);
});
math1('CSC', 'Returns the cosecant of an angle.', (x) => {
  if (x === 0) throw ERR.DIV0;
  return 1 / Math.sin(x);
});
math1('CSCH', 'Returns the hyperbolic cosecant of an angle.', (x) => {
  if (x === 0) throw ERR.DIV0;
  return 1 / Math.sinh(x);
});
math1('SEC', 'Returns the secant of an angle.', (x) => 1 / Math.cos(x));
math1('SECH', 'Returns the hyperbolic secant of an angle.', (x) => 1 / Math.cosh(x));
math1('ACOT', 'Returns the arccotangent of a number.', (x) => Math.PI / 2 - Math.atan(x));
math1('ACOTH', 'Returns the hyperbolic arccotangent of a number.', (x) => {
  check(Math.abs(x) > 1);
  return 0.5 * Math.log((x + 1) / (x - 1));
});
fn('ATAN2', M, 2, 2, 'v', 'x_num, y_num', 'Returns the arctangent from x and y coordinates.', ([a, b]) => {
  const x = num(a);
  const y = num(b);
  if (x === 0 && y === 0) throw ERR.DIV0;
  return Math.atan2(y, x);
});

/* ------------------------------------------------------ combinatorics */

function combin(n: number, k: number): number {
  if (k < 0 || n < 0 || k > n) throw ERR.NUM;
  k = Math.min(k, n - k);
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return Math.round(r);
}

fn('COMBIN', M, 2, 2, 'v', 'number, number_chosen', 'Returns the number of combinations for a given number of objects.', ([a, b]) => combin(Math.trunc(num(a)), Math.trunc(num(b))));
fn('COMBINA', M, 2, 2, 'v', 'number, number_chosen', 'Returns the number of combinations with repetitions.', ([a, b]) => {
  const n = Math.trunc(num(a));
  const k = Math.trunc(num(b));
  check(n >= 0 && k >= 0);
  if (n === 0 && k === 0) return 1;
  return combin(n + k - 1, k);
});
fn('PERMUT', M, 2, 2, 'v', 'number, number_chosen', 'Returns the number of permutations for a given number of objects.', ([a, b]) => {
  const n = Math.trunc(num(a));
  const k = Math.trunc(num(b));
  check(n >= 0 && k >= 0 && k <= n);
  let r = 1;
  for (let i = 0; i < k; i++) r *= n - i;
  return r;
});
fn('PERMUTATIONA', M, 2, 2, 'v', 'number, number_chosen', 'Returns the number of permutations with repetitions.', ([a, b]) => {
  const n = Math.trunc(num(a));
  const k = Math.trunc(num(b));
  check(n >= 0 && k >= 0);
  return Math.pow(n, k);
});
fn('MULTINOMIAL', M, 1, 255, 'r', 'number1, [number2], …', 'Returns the multinomial of a set of numbers.', (args, ctx) => {
  const xs = collectNums(args, ctx).map((x) => Math.trunc(x));
  let total = 0;
  let r = 1;
  for (const x of xs) {
    check(x >= 0);
    for (let i = 1; i <= x; i++) {
      total++;
      r = (r * total) / i;
    }
  }
  return Math.round(r);
});

function gcd2(a: number, b: number): number {
  while (b) [a, b] = [b, a % b];
  return a;
}
fn('GCD', M, 1, 255, 'r', 'number1, [number2], …', 'Returns the greatest common divisor.', (args, ctx) => {
  const xs = collectNums(args, ctx).map((x) => Math.trunc(x));
  for (const x of xs) check(x >= 0);
  return xs.reduce((a, b) => gcd2(a, b), 0);
});
fn('LCM', M, 1, 255, 'r', 'number1, [number2], …', 'Returns the least common multiple.', (args, ctx) => {
  const xs = collectNums(args, ctx).map((x) => Math.trunc(x));
  for (const x of xs) check(x >= 0);
  if (xs.some((x) => x === 0)) return 0;
  return xs.reduce((a, b) => (a * b) / gcd2(a, b), 1);
});

/* ------------------------------------------------------------- random */

fn('RAND', M, 0, 0, 'v', '', 'Returns a random number between 0 and 1.', (_a, ctx: EvalContext) => ctx.random(), { volatile: true });
fn('RANDBETWEEN', M, 2, 2, 'v', 'bottom, top', 'Returns a random integer between the numbers you specify.', ([a, b], ctx: EvalContext) => {
  const lo = Math.ceil(num(a));
  const hi = Math.floor(num(b));
  check(lo <= hi);
  return lo + Math.floor(ctx.random() * (hi - lo + 1));
}, { volatile: true });
fn('RANDARRAY', M, 0, 5, 'v', '[rows], [columns], [min], [max], [whole_number]', 'Returns an array of random numbers.', (args, ctx: EvalContext) => {
  const rows = args[0] == null ? 1 : Math.trunc(num(args[0]));
  const cols = args[1] == null ? 1 : Math.trunc(num(args[1]));
  const lo = args[2] == null ? 0 : num(args[2]);
  const hi = args[3] == null ? 1 : num(args[3]);
  const whole = args[4] == null ? false : bool(args[4]);
  if (rows < 1 || cols < 1) throw ERR.CALC;
  check(lo <= hi, ERR.VALUE);
  const data: Scalar[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) row.push(whole ? Math.ceil(lo) + Math.floor(ctx.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + ctx.random() * (hi - lo));
    data.push(row);
  }
  return matrix(data);
}, { volatile: true });

fn('SEQUENCE', M, 1, 4, 'v', 'rows, [columns], [start], [step]', 'Generates a list of sequential numbers in an array.', (args) => {
  const rows = Math.trunc(num(args[0]));
  const cols = args[1] == null ? 1 : Math.trunc(num(args[1]));
  const start = args[2] == null ? 1 : num(args[2]);
  const step = args[3] == null ? 1 : num(args[3]);
  if (rows < 1 || cols < 1) throw ERR.CALC;
  if (rows * cols > 5_000_000) throw ERR.NUM;
  const data: Scalar[][] = [];
  let v = start;
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) {
      row.push(clean(v));
      v += step;
    }
    data.push(row);
  }
  return matrix(data);
});

/* ------------------------------------------------------ conditional sums */

fn('SUMIF', M, 2, 3, 'rvr', 'range, criteria, [sum_range]', 'Adds the cells specified by a given condition or criteria.', (args: Value[], ctx) => {
  const test = makeCriteria(args[1] as Scalar);
  const g = grid(args[0], ctx);
  const R = g.length;
  const C = g[0]?.length ?? 0;
  const sg = args.length > 2 && args[2] !== null ? resizedGrid(args[2], R, C, ctx) : g;
  const xs: number[] = [];
  for (let r = 0; r < R; r++)
    for (let c = 0; c < C; c++)
      if (test(g[r][c])) {
        const v = sg[r][c];
        if (isErr(v)) throw asErr(v);
        if (typeof v === 'number') xs.push(v);
      }
  return sum(xs);
});

fn('SUMIFS', M, 3, 255, 'r|rv', 'sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], …', 'Adds the cells in a range that meet multiple criteria.', (args: Value[], ctx) => {
  if ((args.length - 1) % 2) throw ERR.VALUE;
  const pairs: [Value, Scalar][] = [];
  for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1] as Scalar]);
  const sg = grid(args[0], ctx);
  const xs: number[] = [];
  const [R, C] = matchIfs(pairs, ctx, (r, c) => {
    const v = sg[r]?.[c];
    if (isErr(v)) throw asErr(v);
    if (typeof v === 'number') xs.push(v);
  });
  if (sg.length !== R || (sg[0]?.length ?? 0) !== C) throw ERR.VALUE;
  return sum(xs);
});

/* ---------------------------------------------------------- SUBTOTAL */

const SUBTOTAL_FNS: Record<number, string> = {
  1: 'AVERAGE',
  2: 'COUNT',
  3: 'COUNTA',
  4: 'MAX',
  5: 'MIN',
  6: 'PRODUCT',
  7: 'STDEV.S',
  8: 'STDEV.P',
  9: 'SUM',
  10: 'VAR.S',
  11: 'VAR.P',
};

fn('SUBTOTAL', M, 2, 255, 'vr', 'function_num, ref1, [ref2], …', 'Returns a subtotal in a list, ignoring other subtotals (101-111 also ignore hidden rows).', (args: Value[], ctx) => {
  const code = Math.trunc(num(args[0] as Scalar));
  const skipHidden = code > 100;
  const name = SUBTOTAL_FNS[skipHidden ? code - 100 : code];
  if (!name) throw ERR.VALUE;
  return aggregateBy(name, args.slice(1), ctx, { skipHidden, skipSubtotals: true, skipErrors: false });
});

const AGGREGATE_FNS: Record<number, string> = {
  ...SUBTOTAL_FNS,
  12: 'MEDIAN',
  13: 'MODE.SNGL',
  14: 'LARGE',
  15: 'SMALL',
  16: 'PERCENTILE.INC',
  17: 'QUARTILE.INC',
  18: 'PERCENTILE.EXC',
  19: 'QUARTILE.EXC',
};

fn('AGGREGATE', M, 3, 255, 'vvr', 'function_num, options, ref1, [ref2], …', 'Returns an aggregate (SUM, AVERAGE, LARGE…) with options to ignore hidden rows and errors.', (args: Value[], ctx) => {
  const code = Math.trunc(num(args[0] as Scalar));
  const opt = Math.trunc(num(args[1] as Scalar));
  const name = AGGREGATE_FNS[code];
  if (!name || opt < 0 || opt > 7) throw ERR.VALUE;
  const skipHidden = opt === 1 || opt === 3 || opt === 5 || opt === 7;
  const skipErrors = opt === 2 || opt === 3 || opt === 6 || opt === 7;
  const skipSubtotals = opt <= 3;
  if (code >= 14) {
    // LARGE/SMALL/PERCENTILE/QUARTILE take (array, k)
    return aggregateBy(name, [args[2]], ctx, { skipHidden, skipErrors, skipSubtotals, k: args[3] as Scalar });
  }
  return aggregateBy(name, args.slice(2), ctx, { skipHidden, skipErrors, skipSubtotals });
});

/* ------------------------------------------------------------ matrices */

function numMatrix(v: Value, ctx: EvalContext): number[][] {
  return grid(v, ctx).map((row) =>
    row.map((x) => {
      if (isErr(x)) throw asErr(x);
      if (typeof x !== 'number') throw ERR.VALUE;
      return x;
    }),
  );
}

fn('MMULT', M, 2, 2, 'r', 'array1, array2', 'Returns the matrix product of two arrays.', ([a, b]: Value[], ctx) => {
  const A = numMatrix(a, ctx);
  const B = numMatrix(b, ctx);
  const n = A.length;
  const m = A[0].length;
  if (B.length !== m) throw ERR.VALUE;
  const p = B[0].length;
  const out: Scalar[][] = [];
  for (let i = 0; i < n; i++) {
    const row: Scalar[] = [];
    for (let j = 0; j < p; j++) {
      let s = 0;
      for (let k = 0; k < m; k++) s += A[i][k] * B[k][j];
      row.push(s);
    }
    out.push(row);
  }
  return matrix(out);
});

function lu(A: number[][]): { LU: number[][]; perm: number[]; sign: number } | null {
  const n = A.length;
  const LU = A.map((r) => r.slice());
  const perm = [...Array(n).keys()];
  let sign = 1;
  for (let k = 0; k < n; k++) {
    let p = k;
    for (let i = k + 1; i < n; i++) if (Math.abs(LU[i][k]) > Math.abs(LU[p][k])) p = i;
    if (LU[p][k] === 0) return null;
    if (p !== k) {
      [LU[p], LU[k]] = [LU[k], LU[p]];
      [perm[p], perm[k]] = [perm[k], perm[p]];
      sign = -sign;
    }
    for (let i = k + 1; i < n; i++) {
      LU[i][k] /= LU[k][k];
      for (let j = k + 1; j < n; j++) LU[i][j] -= LU[i][k] * LU[k][j];
    }
  }
  return { LU, perm, sign };
}

fn('MDETERM', M, 1, 1, 'r', 'array', 'Returns the matrix determinant of an array.', ([a]: Value[], ctx) => {
  const A = numMatrix(a, ctx);
  if (A.length !== A[0].length) throw ERR.VALUE;
  const d = lu(A);
  if (!d) return 0;
  let det = d.sign;
  for (let i = 0; i < A.length; i++) det *= d.LU[i][i];
  return clean(det);
});

fn('MINVERSE', M, 1, 1, 'r', 'array', 'Returns the inverse matrix of an array.', ([a]: Value[], ctx) => {
  const A = numMatrix(a, ctx);
  const n = A.length;
  if (n !== A[0].length) throw ERR.VALUE;
  const aug = A.map((row, i) => [...row, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let k = 0; k < n; k++) {
    let p = k;
    for (let i = k + 1; i < n; i++) if (Math.abs(aug[i][k]) > Math.abs(aug[p][k])) p = i;
    if (Math.abs(aug[p][k]) < 1e-15) throw ERR.NUM;
    [aug[p], aug[k]] = [aug[k], aug[p]];
    const piv = aug[k][k];
    for (let j = 0; j < 2 * n; j++) aug[k][j] /= piv;
    for (let i = 0; i < n; i++) {
      if (i === k) continue;
      const f = aug[i][k];
      if (f) for (let j = 0; j < 2 * n; j++) aug[i][j] -= f * aug[k][j];
    }
  }
  return matrix(aug.map((row) => row.slice(n).map((x) => clean(x))));
});

fn('MUNIT', M, 1, 1, 'v', 'dimension', 'Returns the unit matrix for the specified dimension.', ([a]) => {
  const n = Math.trunc(num(a));
  check(n >= 1, ERR.VALUE);
  return matrix(Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))));
});

function pairSum(name: string, desc: string, f: (x: number, y: number) => number): void {
  fn(name, M, 2, 2, 'r', 'array_x, array_y', desc, ([a, b]: Value[], ctx) => {
    const xa = flat(a, ctx);
    const xb = flat(b, ctx);
    if (xa.length !== xb.length) throw ERR.NA;
    let s = 0;
    for (let i = 0; i < xa.length; i++) {
      const x = xa[i];
      const y = xb[i];
      if (isErr(x)) throw asErr(x);
      if (isErr(y)) throw asErr(y);
      if (typeof x === 'number' && typeof y === 'number') s += f(x, y);
    }
    return s;
  });
}
pairSum('SUMX2MY2', 'Returns the sum of the difference of squares of corresponding values.', (x, y) => x * x - y * y);
pairSum('SUMX2PY2', 'Returns the sum of the sum of squares of corresponding values.', (x, y) => x * x + y * y);
pairSum('SUMXMY2', 'Returns the sum of squares of differences of corresponding values.', (x, y) => (x - y) * (x - y));

fn('SERIESSUM', M, 4, 4, 'vvvr', 'x, n, m, coefficients', 'Returns the sum of a power series.', (args: Value[], ctx) => {
  const x = num(args[0] as Scalar);
  const n = num(args[1] as Scalar);
  const m = num(args[2] as Scalar);
  let s = 0;
  let i = 0;
  eachArg([args[3]], ctx, (v) => {
    if (isErr(v)) throw asErr(v);
    s += num(v) * Math.pow(x, n + i * m);
    i++;
  });
  return s;
});

/* ------------------------------------------------------------ numerals */

const ROMAN: [number, string][] = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];

fn('ROMAN', M, 1, 2, 'v', 'number, [form]', 'Converts an arabic numeral to roman, as text.', (args) => {
  let n = Math.trunc(num(args[0]));
  check(n >= 0 && n <= 3999, ERR.VALUE);
  let out = '';
  for (const [v, s] of ROMAN)
    while (n >= v) {
      out += s;
      n -= v;
    }
  return out;
});

fn('ARABIC', M, 1, 1, 'v', 'text', 'Converts a Roman numeral to Arabic.', ([t]) => {
  let s = str(t).trim().toUpperCase();
  let sign = 1;
  if (s.startsWith('-')) {
    sign = -1;
    s = s.slice(1);
  }
  if (s === '') return 0;
  const val: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let total = 0;
  for (let i = 0; i < s.length; i++) {
    const v = val[s[i]];
    if (!v) throw ERR.VALUE;
    const nx = val[s[i + 1]] ?? 0;
    total += v < nx ? -v : v;
  }
  return sign * total;
});

fn('BASE', M, 2, 3, 'v', 'number, radix, [min_length]', 'Converts a number into a text representation with the given radix (base).', (args) => {
  const n = Math.trunc(num(args[0]));
  const radix = Math.trunc(num(args[1]));
  const len = optInt(args, 2, 0);
  check(n >= 0 && n < 2 ** 53 && radix >= 2 && radix <= 36 && len >= 0);
  return n.toString(radix).toUpperCase().padStart(len, '0');
});

fn('DECIMAL', M, 2, 2, 'v', 'text, radix', 'Converts a text representation of a number in a given base into a decimal number.', ([t, r]) => {
  const radix = Math.trunc(num(r));
  check(radix >= 2 && radix <= 36);
  const s = str(t).trim().toLowerCase();
  if (!s) return 0;
  let v = 0;
  for (const ch of s) {
    const d = parseInt(ch, 36);
    if (Number.isNaN(d) || d >= radix) throw ERR.NUM;
    v = v * radix + d;
  }
  return v;
});

/* -------------------------------------------------------------- gamma */

const LANCZOS = [
  0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905, -0.13857109526572012,
  9.9843695780195716e-6, 1.5056327351493116e-7,
];

export function gamma(x: number): number {
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
  x -= 1;
  let a = LANCZOS[0];
  const t = x + 7.5;
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (x + i);
  return Math.sqrt(2 * Math.PI) * Math.pow(t, x + 0.5) * Math.exp(-t) * a;
}

export function gammaln(x: number): number {
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - gammaln(1 - x);
  x -= 1;
  let a = LANCZOS[0];
  const t = x + 7.5;
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (x + i);
  return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
}
