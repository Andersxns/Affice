import { ERR, isErr, asErr, isRef, makeCriteria, matrix, type Scalar, type Value } from '../values';
import type { EvalContext } from '../evaluator';
import { numEquals } from '../evaluator';
import { check, collectNums, eachArg, eachValue, flat, fn, fullDims, grid, matchIfs, mean, num, pairedNums, resizedGrid, sortedNums, sum, variance } from './helpers';

const S = 'Statistical' as const;

/* ------------------------------------------------------------ helpers */

function percentileInc(xs: number[], k: number): number {
  check(xs.length > 0 && k >= 0 && k <= 1);
  const s = sortedNums(xs);
  const pos = k * (s.length - 1);
  const lo = Math.floor(pos);
  const frac = pos - lo;
  return lo + 1 < s.length ? s[lo] + frac * (s[lo + 1] - s[lo]) : s[lo];
}

function percentileExc(xs: number[], k: number): number {
  const n = xs.length;
  check(n > 0 && k > 0 && k < 1);
  const pos = k * (n + 1) - 1;
  check(pos >= 0 && pos <= n - 1);
  const s = sortedNums(xs);
  const lo = Math.floor(pos);
  const frac = pos - lo;
  return lo + 1 < n ? s[lo] + frac * (s[lo + 1] - s[lo]) : s[lo];
}

function median(xs: number[]): number {
  check(xs.length > 0);
  const s = sortedNums(xs);
  const m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

function modeSngl(xs: number[]): number {
  const counts = new Map<number, number>();
  let best = NaN;
  let bestN = 1;
  for (const x of xs) {
    const n = (counts.get(x) ?? 0) + 1;
    counts.set(x, n);
    if (n > bestN) {
      bestN = n;
      best = x;
    }
  }
  if (Number.isNaN(best)) throw ERR.NA;
  // ties: the value that appears first wins
  for (const x of xs) if (counts.get(x) === bestN) return x;
  return best;
}

function kth(xs: number[], k: number, largest: boolean): number {
  const kk = Math.ceil(k);
  check(xs.length > 0 && kk >= 1 && kk <= xs.length);
  const s = sortedNums(xs);
  return largest ? s[s.length - kk] : s[kk - 1];
}

const STAT: Record<string, (xs: number[], extra: { countA: number; k?: Scalar }) => number> = {
  AVERAGE: (xs) => mean(xs),
  COUNT: (xs) => xs.length,
  COUNTA: (_xs, e) => e.countA,
  MAX: (xs) => (xs.length ? Math.max(...xs) : 0),
  MIN: (xs) => (xs.length ? Math.min(...xs) : 0),
  PRODUCT: (xs) => (xs.length ? xs.reduce((a, b) => a * b, 1) : 0),
  'STDEV.S': (xs) => Math.sqrt(variance(xs, true)),
  'STDEV.P': (xs) => Math.sqrt(variance(xs, false)),
  SUM: (xs) => sum(xs),
  'VAR.S': (xs) => variance(xs, true),
  'VAR.P': (xs) => variance(xs, false),
  MEDIAN: (xs) => median(xs),
  'MODE.SNGL': (xs) => modeSngl(xs),
  LARGE: (xs, e) => kth(xs, num(e.k ?? null), true),
  SMALL: (xs, e) => kth(xs, num(e.k ?? null), false),
  'PERCENTILE.INC': (xs, e) => percentileInc(xs, num(e.k ?? null)),
  'QUARTILE.INC': (xs, e) => quartile(xs, num(e.k ?? null), false),
  'PERCENTILE.EXC': (xs, e) => percentileExc(xs, num(e.k ?? null)),
  'QUARTILE.EXC': (xs, e) => quartile(xs, num(e.k ?? null), true),
};

function quartile(xs: number[], q: number, exc: boolean): number {
  const qq = Math.trunc(q);
  if (exc) {
    check(qq >= 1 && qq <= 3);
    return percentileExc(xs, qq / 4);
  }
  check(qq >= 0 && qq <= 4);
  return percentileInc(xs, qq / 4);
}

const NESTED = /\b(SUBTOTAL|AGGREGATE)\s*\(/i;

/** Shared engine of SUBTOTAL and AGGREGATE. */
export function aggregateBy(
  name: string,
  refs: Value[],
  ctx: EvalContext,
  opts: { skipHidden: boolean; skipErrors: boolean; skipSubtotals: boolean; k?: Scalar },
): number {
  const xs: number[] = [];
  let countA = 0;
  for (const a of refs) {
    if (isRef(a)) {
      const cells: [Scalar, number, number][] = [];
      ctx.each(a, (v, r, c) => cells.push([v, r, c]));
      cells.sort((p, q) => p[1] - q[1] || p[2] - q[2]);
      for (const [v, r, c] of cells) {
        const st = ctx.rowState?.(a.sheet, r);
        if (st === 'filtered' || (st === 'hidden' && opts.skipHidden)) continue;
        if (opts.skipSubtotals) {
          const f = ctx.formulaText(a.sheet, r, c);
          if (f && NESTED.test(f)) continue;
        }
        if (isErr(v)) {
          if (opts.skipErrors) continue;
          throw asErr(v);
        }
        if (v !== null) countA++;
        if (typeof v === 'number') xs.push(v);
      }
    } else {
      eachValue(a, ctx, (v) => {
        if (isErr(v)) {
          if (opts.skipErrors) return;
          throw asErr(v);
        }
        if (v !== null) countA++;
        if (typeof v === 'number') xs.push(v);
      });
    }
  }
  const f = STAT[name];
  if (!f) throw ERR.VALUE;
  return f(xs, { countA, k: opts.k });
}

/* ----------------------------------------------------------- averages */

fn('AVERAGE', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the average (arithmetic mean) of its arguments.', (args, ctx) => mean(collectNums(args, ctx)));
fn('AVERAGEA', S, 1, 255, 'r', 'value1, [value2], …', 'Returns the average of its arguments, counting text as 0 and TRUE as 1.', (args, ctx) => mean(collectNums(args, ctx, 'a')));

fn('AVERAGEIF', S, 2, 3, 'rvr', 'range, criteria, [average_range]', 'Returns the average of the cells that meet a criterion.', (args: Value[], ctx) => {
  const test = makeCriteria(args[1] as Scalar);
  const g = grid(args[0], ctx);
  const R = g.length;
  const C = g[0]?.length ?? 0;
  const ag = args.length > 2 && args[2] !== null ? resizedGrid(args[2], R, C, ctx) : g;
  const xs: number[] = [];
  for (let r = 0; r < R; r++)
    for (let c = 0; c < C; c++)
      if (test(g[r][c])) {
        const v = ag[r][c];
        if (isErr(v)) throw asErr(v);
        if (typeof v === 'number') xs.push(v);
      }
  return mean(xs);
});

function ifsCollect(args: Value[], ctx: EvalContext): number[] {
  if ((args.length - 1) % 2) throw ERR.VALUE;
  const pairs: [Value, Scalar][] = [];
  for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1] as Scalar]);
  const vg = grid(args[0], ctx);
  const xs: number[] = [];
  const [R, C] = matchIfs(pairs, ctx, (r, c) => {
    const v = vg[r]?.[c];
    if (isErr(v)) throw asErr(v);
    if (typeof v === 'number') xs.push(v);
  });
  if (vg.length !== R || (vg[0]?.length ?? 0) !== C) throw ERR.VALUE;
  return xs;
}

fn('AVERAGEIFS', S, 3, 255, 'r|rv', 'average_range, criteria_range1, criteria1, …', 'Returns the average of the cells that meet multiple criteria.', (args, ctx) => mean(ifsCollect(args, ctx)));
fn('MAXIFS', S, 3, 255, 'r|rv', 'max_range, criteria_range1, criteria1, …', 'Returns the maximum among cells specified by a set of conditions.', (args, ctx) => {
  const xs = ifsCollect(args, ctx);
  return xs.length ? Math.max(...xs) : 0;
});
fn('MINIFS', S, 3, 255, 'r|rv', 'min_range, criteria_range1, criteria1, …', 'Returns the minimum among cells specified by a set of conditions.', (args, ctx) => {
  const xs = ifsCollect(args, ctx);
  return xs.length ? Math.min(...xs) : 0;
});

/* ------------------------------------------------------------- counts */

fn('COUNT', S, 1, 255, 'r', 'value1, [value2], …', 'Counts how many cells contain numbers.', (args, ctx) => {
  let n = 0;
  eachArg(args, ctx, (v, direct) => {
    if (typeof v === 'number') n++;
    else if (direct && v !== null && !isErr(v)) {
      if (typeof v === 'boolean') n++;
      else if (typeof v === 'string') {
        try {
          num(v);
          n++;
        } catch {
          /* not a number */
        }
      }
    }
  });
  return n;
});

fn('COUNTA', S, 1, 255, 'r', 'value1, [value2], …', 'Counts how many cells are not empty.', (args, ctx) => {
  let n = 0;
  eachArg(args, ctx, (v, direct) => {
    if (v !== null || direct) n++;
  });
  return n;
});

fn('COUNTBLANK', S, 1, 1, 'r', 'range', 'Counts empty cells in a range.', ([a]: Value[], ctx) => {
  const [R, C] = fullDims(a);
  let filled = 0;
  eachValue(a, ctx, (v) => {
    if (v !== null && v !== '') filled++;
  });
  return R * C - filled;
});

fn('COUNTIF', S, 2, 2, 'rv', 'range, criteria', 'Counts the cells in a range that meet a condition.', ([rg, crit]: [Value, Scalar], ctx) => {
  const test = makeCriteria(crit);
  // Criteria that can match blanks need a dense scan; others only visit filled cells.
  if (test(null)) {
    const g = grid(rg, ctx);
    let n = 0;
    for (const row of g) for (const v of row) if (test(v)) n++;
    if (isRef(rg)) {
      const [R, C] = fullDims(rg);
      n += R * C - g.length * (g[0]?.length ?? 0);
    }
    return n;
  }
  let n = 0;
  eachValue(rg, ctx, (v) => {
    if (test(v)) n++;
  });
  return n;
});

fn('COUNTIFS', S, 2, 254, '|rv', 'criteria_range1, criteria1, [criteria_range2, criteria2], …', 'Counts the cells that meet multiple criteria.', (args: Value[], ctx) => {
  if (args.length % 2) throw ERR.VALUE;
  const pairs: [Value, Scalar][] = [];
  for (let i = 0; i < args.length; i += 2) pairs.push([args[i], args[i + 1] as Scalar]);
  let n = 0;
  matchIfs(pairs, ctx, () => n++);
  return n;
});

/* ------------------------------------------------------------ extremes */

fn('MAX', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the largest value in a set of values.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  return xs.length ? xs.reduce((a, b) => (b > a ? b : a)) : 0;
});
fn('MIN', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the smallest value in a set of values.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  return xs.length ? xs.reduce((a, b) => (b < a ? b : a)) : 0;
});
fn('MAXA', S, 1, 255, 'r', 'value1, [value2], …', 'Returns the largest value, counting text as 0 and TRUE as 1.', (args, ctx) => {
  const xs = collectNums(args, ctx, 'a');
  return xs.length ? Math.max(...xs) : 0;
});
fn('MINA', S, 1, 255, 'r', 'value1, [value2], …', 'Returns the smallest value, counting text as 0 and TRUE as 1.', (args, ctx) => {
  const xs = collectNums(args, ctx, 'a');
  return xs.length ? Math.min(...xs) : 0;
});
fn('MEDIAN', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the median of the given numbers.', (args, ctx) => median(collectNums(args, ctx)));
fn('MODE.SNGL', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the most frequently occurring value.', (args, ctx) => modeSngl(collectNums(args, ctx)));
fn('MODE', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the most frequently occurring value.', (args, ctx) => modeSngl(collectNums(args, ctx)));
fn('MODE.MULT', S, 1, 255, 'r', 'number1, [number2], …', 'Returns a vertical array of the most frequently occurring values.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  const counts = new Map<number, number>();
  for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1);
  const best = Math.max(0, ...counts.values());
  if (best < 2) throw ERR.NA;
  const out: number[] = [];
  for (const x of xs) if (counts.get(x) === best && !out.includes(x)) out.push(x);
  return matrix(out.map((x) => [x]));
});

fn('LARGE', S, 2, 2, 'rv', 'array, k', 'Returns the k-th largest value in a data set.', ([a, k]: [Value, Scalar], ctx) => kth(collectNums([a], ctx), num(k), true));
fn('SMALL', S, 2, 2, 'rv', 'array, k', 'Returns the k-th smallest value in a data set.', ([a, k]: [Value, Scalar], ctx) => kth(collectNums([a], ctx), num(k), false));

/* ---------------------------------------------------------- deviation */

fn('STDEV.S', S, 1, 255, 'r', 'number1, [number2], …', 'Estimates standard deviation based on a sample.', (args, ctx) => Math.sqrt(variance(collectNums(args, ctx), true)));
fn('STDEV', S, 1, 255, 'r', 'number1, [number2], …', 'Estimates standard deviation based on a sample.', (args, ctx) => Math.sqrt(variance(collectNums(args, ctx), true)));
fn('STDEV.P', S, 1, 255, 'r', 'number1, [number2], …', 'Calculates standard deviation based on the entire population.', (args, ctx) => Math.sqrt(variance(collectNums(args, ctx), false)));
fn('STDEVP', S, 1, 255, 'r', 'number1, [number2], …', 'Calculates standard deviation based on the entire population.', (args, ctx) => Math.sqrt(variance(collectNums(args, ctx), false)));
fn('STDEVA', S, 1, 255, 'r', 'value1, [value2], …', 'Estimates standard deviation including text and logical values.', (args, ctx) => Math.sqrt(variance(collectNums(args, ctx, 'a'), true)));
fn('STDEVPA', S, 1, 255, 'r', 'value1, [value2], …', 'Population standard deviation including text and logical values.', (args, ctx) => Math.sqrt(variance(collectNums(args, ctx, 'a'), false)));
fn('VAR.S', S, 1, 255, 'r', 'number1, [number2], …', 'Estimates variance based on a sample.', (args, ctx) => variance(collectNums(args, ctx), true));
fn('VAR', S, 1, 255, 'r', 'number1, [number2], …', 'Estimates variance based on a sample.', (args, ctx) => variance(collectNums(args, ctx), true));
fn('VAR.P', S, 1, 255, 'r', 'number1, [number2], …', 'Calculates variance based on the entire population.', (args, ctx) => variance(collectNums(args, ctx), false));
fn('VARP', S, 1, 255, 'r', 'number1, [number2], …', 'Calculates variance based on the entire population.', (args, ctx) => variance(collectNums(args, ctx), false));
fn('VARA', S, 1, 255, 'r', 'value1, [value2], …', 'Estimates variance including text and logical values.', (args, ctx) => variance(collectNums(args, ctx, 'a'), true));
fn('VARPA', S, 1, 255, 'r', 'value1, [value2], …', 'Population variance including text and logical values.', (args, ctx) => variance(collectNums(args, ctx, 'a'), false));

fn('AVEDEV', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the average of the absolute deviations from their mean.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  const m = mean(xs);
  return sum(xs.map((x) => Math.abs(x - m))) / xs.length;
});
fn('DEVSQ', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the sum of squares of deviations.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  if (!xs.length) throw ERR.NUM;
  const m = mean(xs);
  return sum(xs.map((x) => (x - m) * (x - m)));
});
fn('GEOMEAN', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the geometric mean.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  check(xs.length > 0 && xs.every((x) => x > 0));
  return Math.exp(sum(xs.map(Math.log)) / xs.length);
});
fn('HARMEAN', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the harmonic mean.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  check(xs.length > 0 && xs.every((x) => x > 0));
  return xs.length / sum(xs.map((x) => 1 / x));
});
fn('TRIMMEAN', S, 2, 2, 'rv', 'array, percent', 'Returns the mean of the interior of a data set.', ([a, p]: [Value, Scalar], ctx) => {
  const xs = sortedNums(collectNums([a], ctx));
  const pct = num(p);
  check(pct >= 0 && pct < 1 && xs.length > 0);
  let k = Math.floor((xs.length * pct) / 2);
  if (k * 2 >= xs.length) k = 0;
  return mean(xs.slice(k, xs.length - k));
});
fn('KURT', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the kurtosis of a data set.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  const n = xs.length;
  if (n < 4) throw ERR.DIV0;
  const m = mean(xs);
  const sd = Math.sqrt(variance(xs, true));
  if (sd === 0) throw ERR.DIV0;
  const s4 = sum(xs.map((x) => ((x - m) / sd) ** 4));
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
});
fn('SKEW', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the skewness of a distribution.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  const n = xs.length;
  if (n < 3) throw ERR.DIV0;
  const m = mean(xs);
  const sd = Math.sqrt(variance(xs, true));
  if (sd === 0) throw ERR.DIV0;
  return (n / ((n - 1) * (n - 2))) * sum(xs.map((x) => ((x - m) / sd) ** 3));
});
fn('SKEW.P', S, 1, 255, 'r', 'number1, [number2], …', 'Returns the skewness of a distribution based on a population.', (args, ctx) => {
  const xs = collectNums(args, ctx);
  const n = xs.length;
  if (n < 1) throw ERR.DIV0;
  const m = mean(xs);
  const sd = Math.sqrt(variance(xs, false));
  if (sd === 0) throw ERR.DIV0;
  return sum(xs.map((x) => ((x - m) / sd) ** 3)) / n;
});
fn('STANDARDIZE', S, 3, 3, 'v', 'x, mean, standard_dev', 'Returns a normalized value.', ([x, m, s]) => {
  const sd = num(s);
  check(sd > 0);
  return (num(x) - num(m)) / sd;
});
fn('FISHER', S, 1, 1, 'v', 'x', 'Returns the Fisher transformation.', ([x]) => {
  const v = num(x);
  check(v > -1 && v < 1);
  return 0.5 * Math.log((1 + v) / (1 - v));
});
fn('FISHERINV', S, 1, 1, 'v', 'y', 'Returns the inverse of the Fisher transformation.', ([y]) => {
  const e = Math.exp(2 * num(y));
  return (e - 1) / (e + 1);
});

/* --------------------------------------------------------- percentiles */

fn('PERCENTILE.INC', S, 2, 2, 'rv', 'array, k', 'Returns the k-th percentile of values in a range (0..1 inclusive).', ([a, k]: [Value, Scalar], ctx) => percentileInc(collectNums([a], ctx), num(k)));
fn('PERCENTILE', S, 2, 2, 'rv', 'array, k', 'Returns the k-th percentile of values in a range.', ([a, k]: [Value, Scalar], ctx) => percentileInc(collectNums([a], ctx), num(k)));
fn('PERCENTILE.EXC', S, 2, 2, 'rv', 'array, k', 'Returns the k-th percentile of values in a range (0..1 exclusive).', ([a, k]: [Value, Scalar], ctx) => percentileExc(collectNums([a], ctx), num(k)));
fn('QUARTILE.INC', S, 2, 2, 'rv', 'array, quart', 'Returns the quartile of a data set (0..4).', ([a, q]: [Value, Scalar], ctx) => quartile(collectNums([a], ctx), num(q), false));
fn('QUARTILE', S, 2, 2, 'rv', 'array, quart', 'Returns the quartile of a data set (0..4).', ([a, q]: [Value, Scalar], ctx) => quartile(collectNums([a], ctx), num(q), false));
fn('QUARTILE.EXC', S, 2, 2, 'rv', 'array, quart', 'Returns the quartile of a data set (1..3, exclusive).', ([a, q]: [Value, Scalar], ctx) => quartile(collectNums([a], ctx), num(q), true));

function percentRank(xs: number[], x: number, sig: number, exc: boolean): number {
  check(xs.length > 0 && sig >= 1);
  const s = sortedNums(xs);
  const n = s.length;
  if (x < s[0] || x > s[n - 1]) throw ERR.NA;
  let rank: number;
  const idx = s.findIndex((v) => v >= x);
  if (s[idx] === x) {
    rank = idx;
  } else {
    const lo = s[idx - 1];
    const hi = s[idx];
    rank = idx - 1 + (x - lo) / (hi - lo);
  }
  const pr = exc ? (rank + 1) / (n + 1) : n === 1 ? 1 : rank / (n - 1);
  const f = Math.pow(10, sig);
  return Math.floor(pr * f + 1e-9) / f;
}

fn('PERCENTRANK.INC', S, 2, 3, 'rvv', 'array, x, [significance]', 'Returns the rank of a value in a data set as a percentage (inclusive).', (args: Value[], ctx) =>
  percentRank(collectNums([args[0]], ctx), num(args[1] as Scalar), args.length > 2 ? Math.trunc(num(args[2] as Scalar)) : 3, false),
);
fn('PERCENTRANK', S, 2, 3, 'rvv', 'array, x, [significance]', 'Returns the rank of a value in a data set as a percentage.', (args: Value[], ctx) =>
  percentRank(collectNums([args[0]], ctx), num(args[1] as Scalar), args.length > 2 ? Math.trunc(num(args[2] as Scalar)) : 3, false),
);
fn('PERCENTRANK.EXC', S, 2, 3, 'rvv', 'array, x, [significance]', 'Returns the rank of a value in a data set as a percentage (exclusive).', (args: Value[], ctx) =>
  percentRank(collectNums([args[0]], ctx), num(args[1] as Scalar), args.length > 2 ? Math.trunc(num(args[2] as Scalar)) : 3, true),
);

function rank(x: number, xs: number[], asc: boolean, avg: boolean): number {
  let better = 0;
  let same = 0;
  for (const v of xs) {
    if (numEquals(v, x)) same++;
    else if (asc ? v < x : v > x) better++;
  }
  if (!same) throw ERR.NA;
  return avg ? better + (same + 1) / 2 : better + 1;
}

fn('RANK.EQ', S, 2, 3, 'vrv', 'number, ref, [order]', 'Returns the rank of a number in a list (ties share the top rank).', (args: Value[], ctx) =>
  rank(num(args[0] as Scalar), collectNums([args[1]], ctx), args.length > 2 && num(args[2] as Scalar) !== 0, false),
);
fn('RANK', S, 2, 3, 'vrv', 'number, ref, [order]', 'Returns the rank of a number in a list.', (args: Value[], ctx) =>
  rank(num(args[0] as Scalar), collectNums([args[1]], ctx), args.length > 2 && num(args[2] as Scalar) !== 0, false),
);
fn('RANK.AVG', S, 2, 3, 'vrv', 'number, ref, [order]', 'Returns the rank of a number in a list (ties get the average rank).', (args: Value[], ctx) =>
  rank(num(args[0] as Scalar), collectNums([args[1]], ctx), args.length > 2 && num(args[2] as Scalar) !== 0, true),
);

fn('FREQUENCY', S, 2, 2, 'r', 'data_array, bins_array', 'Calculates how often values occur within ranges and returns a vertical array.', ([d, b]: Value[], ctx) => {
  const data = collectNums([d], ctx);
  const bins = collectNums([b], ctx);
  const order = bins.map((v, i) => [v, i] as const).sort((p, q) => p[0] - q[0]);
  const counts = new Array(bins.length + 1).fill(0);
  for (const x of data) {
    let placed = false;
    for (const [v, i] of order) {
      if (x <= v) {
        counts[i]++;
        placed = true;
        break;
      }
    }
    if (!placed) counts[bins.length]++;
  }
  return matrix(counts.map((c) => [c]));
});

/* -------------------------------------------------------- correlation */

function covariance(xs: number[], ys: number[], sample: boolean): number {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) throw ERR.DIV0;
  const mx = mean(xs);
  const my = mean(ys);
  let s = 0;
  for (let i = 0; i < n; i++) s += (xs[i] - mx) * (ys[i] - my);
  return s / (sample ? n - 1 : n);
}

function correl(xs: number[], ys: number[]): number {
  if (xs.length < 2) throw ERR.DIV0;
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < xs.length; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
    syy += (ys[i] - my) ** 2;
  }
  if (sxx === 0 || syy === 0) throw ERR.DIV0;
  return sxy / Math.sqrt(sxx * syy);
}

function regression(knownY: Value, knownX: Value, ctx: EvalContext): { slope: number; intercept: number; xs: number[]; ys: number[] } {
  const [ys, xs] = pairedNums(knownY, knownX, ctx);
  if (xs.length < 1) throw ERR.DIV0;
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < xs.length; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
  }
  if (sxx === 0) throw ERR.DIV0;
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx, xs, ys };
}

fn('CORREL', S, 2, 2, 'r', 'array1, array2', 'Returns the correlation coefficient between two data sets.', ([a, b]: Value[], ctx) => correl(...pairedNums(a, b, ctx)));
fn('PEARSON', S, 2, 2, 'r', 'array1, array2', 'Returns the Pearson product moment correlation coefficient.', ([a, b]: Value[], ctx) => correl(...pairedNums(a, b, ctx)));
fn('RSQ', S, 2, 2, 'r', "known_y's, known_x's", 'Returns the square of the Pearson correlation coefficient.', ([a, b]: Value[], ctx) => correl(...pairedNums(a, b, ctx)) ** 2);
fn('COVARIANCE.P', S, 2, 2, 'r', 'array1, array2', 'Returns population covariance.', ([a, b]: Value[], ctx) => covariance(...pairedNums(a, b, ctx), false));
fn('COVAR', S, 2, 2, 'r', 'array1, array2', 'Returns population covariance.', ([a, b]: Value[], ctx) => covariance(...pairedNums(a, b, ctx), false));
fn('COVARIANCE.S', S, 2, 2, 'r', 'array1, array2', 'Returns sample covariance.', ([a, b]: Value[], ctx) => covariance(...pairedNums(a, b, ctx), true));
fn('SLOPE', S, 2, 2, 'r', "known_y's, known_x's", 'Returns the slope of the linear regression line.', ([y, x]: Value[], ctx) => regression(y, x, ctx).slope);
fn('INTERCEPT', S, 2, 2, 'r', "known_y's, known_x's", 'Returns the intercept of the linear regression line.', ([y, x]: Value[], ctx) => regression(y, x, ctx).intercept);
fn('STEYX', S, 2, 2, 'r', "known_y's, known_x's", 'Returns the standard error of the predicted y-value for each x in the regression.', ([y, x]: Value[], ctx) => {
  const { slope, intercept, xs, ys } = regression(y, x, ctx);
  if (xs.length < 3) throw ERR.DIV0;
  let ss = 0;
  for (let i = 0; i < xs.length; i++) ss += (ys[i] - (intercept + slope * xs[i])) ** 2;
  return Math.sqrt(ss / (xs.length - 2));
});
fn('FORECAST.LINEAR', S, 3, 3, 'vrr', "x, known_y's, known_x's", 'Predicts a future value along a linear trend.', ([x, y, xs]: Value[], ctx) => {
  const { slope, intercept } = regression(y, xs, ctx);
  return intercept + slope * num(x as Scalar);
});
fn('FORECAST', S, 3, 3, 'vrr', "x, known_y's, known_x's", 'Predicts a future value along a linear trend.', ([x, y, xs]: Value[], ctx) => {
  const { slope, intercept } = regression(y, xs, ctx);
  return intercept + slope * num(x as Scalar);
});

function trendValues(args: Value[], ctx: EvalContext, log: boolean): Value {
  const yg = grid(args[0], ctx);
  const ysRaw = flat(args[0], ctx);
  const ys = ysRaw.map((v) => {
    if (typeof v !== 'number') throw ERR.VALUE;
    if (log && v <= 0) throw ERR.NUM;
    return log ? Math.log(v) : v;
  });
  const xs = args.length > 1 && args[1] !== null ? flat(args[1], ctx).map((v) => num(v)) : ys.map((_, i) => i + 1);
  if (xs.length !== ys.length) throw ERR.REF;
  const constant = args.length > 3 && args[3] !== null ? !!num(args[3] as Scalar) : true;
  const mx = constant ? mean(xs) : 0;
  const my = constant ? mean(ys) : 0;
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < xs.length; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
  }
  if (sxx === 0) throw ERR.DIV0;
  const slope = sxy / sxx;
  const icpt = my - slope * mx;
  const newX = args.length > 2 && args[2] !== null ? grid(args[2], ctx) : args.length > 1 && args[1] !== null ? grid(args[1], ctx) : yg.map((row, i) => row.map((_, j) => i * row.length + j + 1));
  return matrix(newX.map((row) => row.map((v) => {
    const y = icpt + slope * num(v);
    return log ? Math.exp(y) : y;
  })));
}

fn('TREND', S, 1, 4, 'r', "known_y's, [known_x's], [new_x's], [const]", 'Returns values along a linear trend.', (args, ctx) => trendValues(args, ctx, false));
fn('GROWTH', S, 1, 4, 'r', "known_y's, [known_x's], [new_x's], [const]", 'Returns values along an exponential trend.', (args, ctx) => trendValues(args, ctx, true));

fn('LINEST', S, 1, 4, 'r', "known_y's, [known_x's], [const], [stats]", 'Returns the parameters of a linear trend (slope and intercept).', (args: Value[], ctx) => {
  const ys = flat(args[0], ctx).map((v) => num(v));
  const xs = args.length > 1 && args[1] !== null ? flat(args[1], ctx).map((v) => num(v)) : ys.map((_, i) => i + 1);
  if (xs.length !== ys.length) throw ERR.REF;
  const constant = args.length > 2 && args[2] !== null ? !!num(args[2] as Scalar) : true;
  const stats = args.length > 3 && args[3] !== null ? !!num(args[3] as Scalar) : false;
  const n = xs.length;
  const mx = constant ? mean(xs) : 0;
  const my = constant ? mean(ys) : 0;
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
  }
  if (sxx === 0) throw ERR.DIV0;
  const m = sxy / sxx;
  const b = constant ? my - m * mx : 0;
  if (!stats) return matrix([[m, b]]);
  const df = n - (constant ? 2 : 1);
  let ssres = 0;
  let sstot = 0;
  for (let i = 0; i < n; i++) {
    ssres += (ys[i] - (m * xs[i] + b)) ** 2;
    sstot += (ys[i] - (constant ? my : 0)) ** 2;
  }
  const sey = Math.sqrt(ssres / df);
  const sem = sey / Math.sqrt(sxx);
  const seb = constant ? sey * Math.sqrt(1 / n + (mx * mx) / sxx) : ERR.NA;
  const r2 = 1 - ssres / sstot;
  const ssreg = sstot - ssres;
  const F = ssreg / 1 / (ssres / df);
  return matrix([
    [m, b],
    [sem, seb],
    [r2, sey],
    [F, df],
    [ssreg, ssres],
  ]);
});
