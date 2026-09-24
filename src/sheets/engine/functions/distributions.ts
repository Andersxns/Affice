import jStat from 'jstat';
import { ERR, type Scalar, type Value } from '../values';
import type { EvalContext } from '../evaluator';
import { bool, check, collectNums, fn, mean, num, optNum, pairedNums, variance } from './helpers';
import { alias } from './registry';

const S = 'Statistical' as const;

const ok = (x: number) => {
  if (!Number.isFinite(x)) throw ERR.NUM;
  return x;
};

/* ------------------------------------------------------------- normal */

function normDist(x: number, m: number, s: number, cum: boolean): number {
  check(s > 0);
  return ok(cum ? jStat.normal.cdf(x, m, s) : jStat.normal.pdf(x, m, s));
}
function normInv(p: number, m: number, s: number): number {
  check(p > 0 && p < 1 && s > 0);
  return ok(jStat.normal.inv(p, m, s));
}

fn('NORM.DIST', S, 4, 4, 'v', 'x, mean, standard_dev, cumulative', 'Returns the normal distribution for the specified mean and standard deviation.', ([x, m, s, c]) => normDist(num(x), num(m), num(s), bool(c)));
fn('NORMDIST', S, 4, 4, 'v', 'x, mean, standard_dev, cumulative', 'Returns the normal cumulative distribution.', ([x, m, s, c]) => normDist(num(x), num(m), num(s), bool(c)));
fn('NORM.INV', S, 3, 3, 'v', 'probability, mean, standard_dev', 'Returns the inverse of the normal cumulative distribution.', ([p, m, s]) => normInv(num(p), num(m), num(s)));
fn('NORMINV', S, 3, 3, 'v', 'probability, mean, standard_dev', 'Returns the inverse of the normal cumulative distribution.', ([p, m, s]) => normInv(num(p), num(m), num(s)));
fn('NORM.S.DIST', S, 2, 2, 'v', 'z, cumulative', 'Returns the standard normal distribution.', ([z, c]) => normDist(num(z), 0, 1, bool(c)));
fn('NORMSDIST', S, 1, 1, 'v', 'z', 'Returns the standard normal cumulative distribution.', ([z]) => normDist(num(z), 0, 1, true));
fn('NORM.S.INV', S, 1, 1, 'v', 'probability', 'Returns the inverse of the standard normal cumulative distribution.', ([p]) => normInv(num(p), 0, 1));
fn('NORMSINV', S, 1, 1, 'v', 'probability', 'Returns the inverse of the standard normal cumulative distribution.', ([p]) => normInv(num(p), 0, 1));
fn('PHI', S, 1, 1, 'v', 'x', 'Returns the value of the density function for a standard normal distribution.', ([x]) => jStat.normal.pdf(num(x), 0, 1));
fn('GAUSS', S, 1, 1, 'v', 'z', 'Returns 0.5 less than the standard normal cumulative distribution.', ([z]) => jStat.normal.cdf(num(z), 0, 1) - 0.5);

function lognormDist(x: number, m: number, s: number, cum: boolean): number {
  check(x > 0 && s > 0);
  return ok(cum ? jStat.lognormal.cdf(x, m, s) : jStat.lognormal.pdf(x, m, s));
}
fn('LOGNORM.DIST', S, 4, 4, 'v', 'x, mean, standard_dev, cumulative', 'Returns the lognormal distribution of x.', ([x, m, s, c]) => lognormDist(num(x), num(m), num(s), bool(c)));
fn('LOGNORMDIST', S, 3, 3, 'v', 'x, mean, standard_dev', 'Returns the cumulative lognormal distribution of x.', ([x, m, s]) => lognormDist(num(x), num(m), num(s), true));
fn('LOGNORM.INV', S, 3, 3, 'v', 'probability, mean, standard_dev', 'Returns the inverse of the lognormal cumulative distribution.', ([p, m, s]) => {
  const pp = num(p);
  check(pp > 0 && pp < 1 && num(s) > 0);
  return ok(jStat.lognormal.inv(pp, num(m), num(s)));
});
alias('LOGINV', 'LOGNORM.INV');

/* ------------------------------------------------------------ student t */

fn('T.DIST', S, 3, 3, 'v', 'x, deg_freedom, cumulative', "Returns the left-tailed Student's t-distribution.", ([x, d, c]) => {
  const df = Math.trunc(num(d));
  check(df >= 1);
  return ok(bool(c) ? jStat.studentt.cdf(num(x), df) : jStat.studentt.pdf(num(x), df));
});
fn('T.DIST.RT', S, 2, 2, 'v', 'x, deg_freedom', "Returns the right-tailed Student's t-distribution.", ([x, d]) => {
  const df = Math.trunc(num(d));
  check(df >= 1);
  return ok(1 - jStat.studentt.cdf(num(x), df));
});
fn('T.DIST.2T', S, 2, 2, 'v', 'x, deg_freedom', "Returns the two-tailed Student's t-distribution.", ([x, d]) => {
  const df = Math.trunc(num(d));
  const xv = num(x);
  check(df >= 1 && xv >= 0);
  return ok(2 * (1 - jStat.studentt.cdf(xv, df)));
});
fn('TDIST', S, 3, 3, 'v', 'x, deg_freedom, tails', "Returns the Student's t-distribution.", ([x, d, t]) => {
  const df = Math.trunc(num(d));
  const xv = num(x);
  const tails = Math.trunc(num(t));
  check(df >= 1 && xv >= 0 && (tails === 1 || tails === 2));
  return ok(tails * (1 - jStat.studentt.cdf(xv, df)));
});
fn('T.INV', S, 2, 2, 'v', 'probability, deg_freedom', "Returns the left-tailed inverse of the Student's t-distribution.", ([p, d]) => {
  const pp = num(p);
  const df = Math.trunc(num(d));
  check(pp > 0 && pp < 1 && df >= 1);
  return ok(jStat.studentt.inv(pp, df));
});
fn('T.INV.2T', S, 2, 2, 'v', 'probability, deg_freedom', "Returns the two-tailed inverse of the Student's t-distribution.", ([p, d]) => {
  const pp = num(p);
  const df = Math.trunc(num(d));
  check(pp > 0 && pp <= 1 && df >= 1);
  return ok(Math.abs(jStat.studentt.inv(pp / 2, df)));
});
alias('TINV', 'T.INV.2T');

function tTest(a: number[], b: number[], tails: number, type: number): number {
  check(tails === 1 || tails === 2);
  let t: number;
  let df: number;
  if (type === 1) {
    if (a.length !== b.length) throw ERR.NA;
    const d = a.map((x, i) => x - b[i]);
    const md = mean(d);
    const sd = Math.sqrt(variance(d, true));
    t = md / (sd / Math.sqrt(d.length));
    df = d.length - 1;
  } else if (type === 2) {
    const n1 = a.length;
    const n2 = b.length;
    const sp = ((n1 - 1) * variance(a, true) + (n2 - 1) * variance(b, true)) / (n1 + n2 - 2);
    t = (mean(a) - mean(b)) / Math.sqrt(sp * (1 / n1 + 1 / n2));
    df = n1 + n2 - 2;
  } else if (type === 3) {
    const v1 = variance(a, true) / a.length;
    const v2 = variance(b, true) / b.length;
    t = (mean(a) - mean(b)) / Math.sqrt(v1 + v2);
    df = (v1 + v2) ** 2 / (v1 ** 2 / (a.length - 1) + v2 ** 2 / (b.length - 1));
  } else throw ERR.NUM;
  const p = 1 - jStat.studentt.cdf(Math.abs(t), df);
  return ok(tails * p);
}

fn('T.TEST', S, 4, 4, 'rrvv', 'array1, array2, tails, type', "Returns the probability associated with a Student's t-test.", (args: Value[], ctx: EvalContext) =>
  tTest(collectNums([args[0]], ctx), collectNums([args[1]], ctx), Math.trunc(num(args[2] as Scalar)), Math.trunc(num(args[3] as Scalar))),
);
alias('TTEST', 'T.TEST');

/* ----------------------------------------------------------- chi square */

fn('CHISQ.DIST', S, 3, 3, 'v', 'x, deg_freedom, cumulative', 'Returns the left-tailed probability of the chi-squared distribution.', ([x, d, c]) => {
  const xv = num(x);
  const df = Math.trunc(num(d));
  check(xv >= 0 && df >= 1);
  return ok(bool(c) ? jStat.chisquare.cdf(xv, df) : jStat.chisquare.pdf(xv, df));
});
fn('CHISQ.DIST.RT', S, 2, 2, 'v', 'x, deg_freedom', 'Returns the right-tailed probability of the chi-squared distribution.', ([x, d]) => {
  const xv = num(x);
  const df = Math.trunc(num(d));
  check(xv >= 0 && df >= 1);
  return ok(1 - jStat.chisquare.cdf(xv, df));
});
alias('CHIDIST', 'CHISQ.DIST.RT');
fn('CHISQ.INV', S, 2, 2, 'v', 'probability, deg_freedom', 'Returns the inverse of the left-tailed probability of the chi-squared distribution.', ([p, d]) => {
  const pp = num(p);
  const df = Math.trunc(num(d));
  check(pp >= 0 && pp < 1 && df >= 1);
  return ok(jStat.chisquare.inv(pp, df));
});
fn('CHISQ.INV.RT', S, 2, 2, 'v', 'probability, deg_freedom', 'Returns the inverse of the right-tailed probability of the chi-squared distribution.', ([p, d]) => {
  const pp = num(p);
  const df = Math.trunc(num(d));
  check(pp > 0 && pp <= 1 && df >= 1);
  return ok(jStat.chisquare.inv(1 - pp, df));
});
alias('CHIINV', 'CHISQ.INV.RT');
fn('CHISQ.TEST', S, 2, 2, 'r', 'actual_range, expected_range', 'Returns the test for independence (chi-squared).', ([a, e]: Value[], ctx: EvalContext) => {
  const [act, exp] = pairedNums(a, e, ctx);
  let chi = 0;
  for (let i = 0; i < act.length; i++) {
    if (exp[i] === 0) throw ERR.DIV0;
    chi += (act[i] - exp[i]) ** 2 / exp[i];
  }
  const g = valueDims(a);
  const df = g[0] > 1 && g[1] > 1 ? (g[0] - 1) * (g[1] - 1) : Math.max(g[0], g[1]) - 1;
  return ok(1 - jStat.chisquare.cdf(chi, Math.max(1, df)));
});
alias('CHITEST', 'CHISQ.TEST');

function valueDims(v: Value): [number, number] {
  if (v && typeof v === 'object' && 'kind' in v) {
    if (v.kind === 'ref') return [v.r2 - v.r1 + 1, v.c2 - v.c1 + 1];
    if (v.kind === 'matrix') return [v.data.length, v.data[0]?.length ?? 0];
  }
  return [1, 1];
}

/* --------------------------------------------------------------- F test */

fn('F.DIST', S, 4, 4, 'v', 'x, deg_freedom1, deg_freedom2, cumulative', 'Returns the (left-tailed) F probability distribution.', ([x, a, b, c]) => {
  const xv = num(x);
  const d1 = Math.trunc(num(a));
  const d2 = Math.trunc(num(b));
  check(xv >= 0 && d1 >= 1 && d2 >= 1);
  return ok(bool(c) ? jStat.centralF.cdf(xv, d1, d2) : jStat.centralF.pdf(xv, d1, d2));
});
fn('F.DIST.RT', S, 3, 3, 'v', 'x, deg_freedom1, deg_freedom2', 'Returns the (right-tailed) F probability distribution.', ([x, a, b]) => {
  const xv = num(x);
  const d1 = Math.trunc(num(a));
  const d2 = Math.trunc(num(b));
  check(xv >= 0 && d1 >= 1 && d2 >= 1);
  return ok(1 - jStat.centralF.cdf(xv, d1, d2));
});
alias('FDIST', 'F.DIST.RT');
fn('F.INV', S, 3, 3, 'v', 'probability, deg_freedom1, deg_freedom2', 'Returns the inverse of the (left-tailed) F probability distribution.', ([p, a, b]) => {
  const pp = num(p);
  check(pp >= 0 && pp <= 1);
  return ok(jStat.centralF.inv(pp, Math.trunc(num(a)), Math.trunc(num(b))));
});
fn('F.INV.RT', S, 3, 3, 'v', 'probability, deg_freedom1, deg_freedom2', 'Returns the inverse of the (right-tailed) F probability distribution.', ([p, a, b]) => {
  const pp = num(p);
  check(pp >= 0 && pp <= 1);
  return ok(jStat.centralF.inv(1 - pp, Math.trunc(num(a)), Math.trunc(num(b))));
});
alias('FINV', 'F.INV.RT');
fn('F.TEST', S, 2, 2, 'r', 'array1, array2', 'Returns the result of an F-test (two-tailed probability that the variances are not significantly different).', ([a, b]: Value[], ctx: EvalContext) => {
  const x = collectNums([a], ctx);
  const y = collectNums([b], ctx);
  const v1 = variance(x, true);
  const v2 = variance(y, true);
  if (v2 === 0) throw ERR.DIV0;
  const f = v1 / v2;
  const p = jStat.centralF.cdf(f, x.length - 1, y.length - 1);
  return ok(2 * Math.min(p, 1 - p));
});
alias('FTEST', 'F.TEST');

/* ------------------------------------------------------ discrete dists */

fn('BINOM.DIST', S, 4, 4, 'v', 'number_s, trials, probability_s, cumulative', 'Returns the individual term binomial distribution probability.', ([k, n, p, c]) => {
  const kk = Math.trunc(num(k));
  const nn = Math.trunc(num(n));
  const pp = num(p);
  check(kk >= 0 && kk <= nn && pp >= 0 && pp <= 1);
  return ok(bool(c) ? jStat.binomial.cdf(kk, nn, pp) : jStat.binomial.pdf(kk, nn, pp));
});
alias('BINOMDIST', 'BINOM.DIST');
fn('BINOM.DIST.RANGE', S, 3, 4, 'v', 'trials, probability_s, number_s, [number_s2]', 'Returns the probability of a trial result using a binomial distribution.', (args) => {
  const n = Math.trunc(num(args[0]));
  const p = num(args[1]);
  const s1 = Math.trunc(num(args[2]));
  const s2 = args.length > 3 && args[3] !== null ? Math.trunc(num(args[3])) : s1;
  check(p >= 0 && p <= 1 && s1 >= 0 && s1 <= n && s2 >= s1 && s2 <= n);
  let sum = 0;
  for (let k = s1; k <= s2; k++) sum += jStat.binomial.pdf(k, n, p);
  return sum;
});
fn('BINOM.INV', S, 3, 3, 'v', 'trials, probability_s, alpha', 'Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.', ([n, p, a]) => {
  const nn = Math.trunc(num(n));
  const pp = num(p);
  const alpha = num(a);
  check(nn >= 0 && pp >= 0 && pp <= 1 && alpha >= 0 && alpha <= 1);
  for (let k = 0; k <= nn; k++) if (jStat.binomial.cdf(k, nn, pp) >= alpha) return k;
  return nn;
});
alias('CRITBINOM', 'BINOM.INV');
fn('NEGBINOM.DIST', S, 4, 4, 'v', 'number_f, number_s, probability_s, cumulative', 'Returns the negative binomial distribution.', ([f, s, p, c]) => {
  const ff = Math.trunc(num(f));
  const ss = Math.trunc(num(s));
  const pp = num(p);
  check(ff >= 0 && ss >= 1 && pp >= 0 && pp <= 1);
  return ok(bool(c) ? jStat.negbin.cdf(ff, ss, pp) : jStat.negbin.pdf(ff, ss, pp));
});
fn('NEGBINOMDIST', S, 3, 3, 'v', 'number_f, number_s, probability_s', 'Returns the negative binomial distribution.', ([f, s, p]) => ok(jStat.negbin.pdf(Math.trunc(num(f)), Math.trunc(num(s)), num(p))));
fn('POISSON.DIST', S, 3, 3, 'v', 'x, mean, cumulative', 'Returns the Poisson distribution.', ([x, m, c]) => {
  const k = Math.trunc(num(x));
  const l = num(m);
  check(k >= 0 && l >= 0);
  return ok(bool(c) ? jStat.poisson.cdf(k, l) : jStat.poisson.pdf(k, l));
});
alias('POISSON', 'POISSON.DIST');
fn('HYPGEOM.DIST', S, 5, 5, 'v', 'sample_s, number_sample, population_s, number_pop, cumulative', 'Returns the hypergeometric distribution.', ([s, n, M, N, c]) => {
  const k = Math.trunc(num(s));
  const nn = Math.trunc(num(n));
  const m = Math.trunc(num(M));
  const NN = Math.trunc(num(N));
  check(k >= 0 && nn > 0 && m > 0 && NN > 0 && nn <= NN && m <= NN);
  if (!bool(c)) return ok(jStat.hypgeom.pdf(k, NN, m, nn));
  let sum = 0;
  for (let i = 0; i <= k; i++) sum += jStat.hypgeom.pdf(i, NN, m, nn);
  return sum;
});
fn('HYPGEOMDIST', S, 4, 4, 'v', 'sample_s, number_sample, population_s, number_pop', 'Returns the hypergeometric distribution.', ([s, n, M, N]) =>
  ok(jStat.hypgeom.pdf(Math.trunc(num(s)), Math.trunc(num(N)), Math.trunc(num(M)), Math.trunc(num(n)))),
);

/* ---------------------------------------------------- continuous dists */

fn('EXPON.DIST', S, 3, 3, 'v', 'x, lambda, cumulative', 'Returns the exponential distribution.', ([x, l, c]) => {
  const xv = num(x);
  const lam = num(l);
  check(xv >= 0 && lam > 0);
  return bool(c) ? 1 - Math.exp(-lam * xv) : lam * Math.exp(-lam * xv);
});
alias('EXPONDIST', 'EXPON.DIST');
fn('GAMMA.DIST', S, 4, 4, 'v', 'x, alpha, beta, cumulative', 'Returns the gamma distribution.', ([x, a, b, c]) => {
  const xv = num(x);
  const al = num(a);
  const be = num(b);
  check(xv >= 0 && al > 0 && be > 0);
  return ok(bool(c) ? jStat.gamma.cdf(xv, al, be) : jStat.gamma.pdf(xv, al, be));
});
alias('GAMMADIST', 'GAMMA.DIST');
fn('GAMMA.INV', S, 3, 3, 'v', 'probability, alpha, beta', 'Returns the inverse of the gamma cumulative distribution.', ([p, a, b]) => {
  const pp = num(p);
  check(pp >= 0 && pp < 1 && num(a) > 0 && num(b) > 0);
  return ok(jStat.gamma.inv(pp, num(a), num(b)));
});
alias('GAMMAINV', 'GAMMA.INV');
fn('BETA.DIST', S, 4, 6, 'v', 'x, alpha, beta, cumulative, [A], [B]', 'Returns the beta probability distribution function.', (args) => {
  const lo = optNum(args, 4, 0);
  const hi = optNum(args, 5, 1);
  const x = num(args[0]);
  const a = num(args[1]);
  const b = num(args[2]);
  check(a > 0 && b > 0 && lo < hi && x >= lo && x <= hi);
  const z = (x - lo) / (hi - lo);
  return ok(bool(args[3]) ? jStat.beta.cdf(z, a, b) : jStat.beta.pdf(z, a, b) / (hi - lo));
});
fn('BETADIST', S, 3, 5, 'v', 'x, alpha, beta, [A], [B]', 'Returns the cumulative beta probability density function.', (args) => {
  const lo = optNum(args, 3, 0);
  const hi = optNum(args, 4, 1);
  const x = num(args[0]);
  check(num(args[1]) > 0 && num(args[2]) > 0 && lo < hi && x >= lo && x <= hi);
  return ok(jStat.beta.cdf((x - lo) / (hi - lo), num(args[1]), num(args[2])));
});
fn('BETA.INV', S, 3, 5, 'v', 'probability, alpha, beta, [A], [B]', 'Returns the inverse of the cumulative beta probability density function.', (args) => {
  const lo = optNum(args, 3, 0);
  const hi = optNum(args, 4, 1);
  const p = num(args[0]);
  check(p > 0 && p < 1 && num(args[1]) > 0 && num(args[2]) > 0 && lo < hi);
  return ok(lo + (hi - lo) * jStat.beta.inv(p, num(args[1]), num(args[2])));
});
alias('BETAINV', 'BETA.INV');
fn('WEIBULL.DIST', S, 4, 4, 'v', 'x, alpha, beta, cumulative', 'Returns the Weibull distribution.', ([x, a, b, c]) => {
  const xv = num(x);
  const al = num(a);
  const be = num(b);
  check(xv >= 0 && al > 0 && be > 0);
  return bool(c) ? 1 - Math.exp(-Math.pow(xv / be, al)) : (al / Math.pow(be, al)) * Math.pow(xv, al - 1) * Math.exp(-Math.pow(xv / be, al));
});
alias('WEIBULL', 'WEIBULL.DIST');

/* ------------------------------------------------------------- testing */

fn('CONFIDENCE.NORM', S, 3, 3, 'v', 'alpha, standard_dev, size', 'Returns the confidence interval for a population mean, using a normal distribution.', ([a, s, n]) => {
  const alpha = num(a);
  const sd = num(s);
  const size = Math.trunc(num(n));
  check(alpha > 0 && alpha < 1 && sd > 0 && size >= 1);
  return (jStat.normal.inv(1 - alpha / 2, 0, 1) * sd) / Math.sqrt(size);
});
alias('CONFIDENCE', 'CONFIDENCE.NORM');
fn('CONFIDENCE.T', S, 3, 3, 'v', 'alpha, standard_dev, size', "Returns the confidence interval for a population mean, using a Student's t distribution.", ([a, s, n]) => {
  const alpha = num(a);
  const sd = num(s);
  const size = Math.trunc(num(n));
  check(alpha > 0 && alpha < 1 && sd > 0);
  if (size < 2) throw ERR.DIV0;
  return (Math.abs(jStat.studentt.inv(alpha / 2, size - 1)) * sd) / Math.sqrt(size);
});
fn('Z.TEST', S, 2, 3, 'rvv', 'array, x, [sigma]', 'Returns the one-tailed P-value of a z-test.', (args: Value[], ctx: EvalContext) => {
  const xs = collectNums([args[0]], ctx);
  const x = num(args[1] as Scalar);
  const sigma = args.length > 2 && args[2] !== null ? num(args[2] as Scalar) : Math.sqrt(variance(xs, true));
  return 1 - jStat.normal.cdf((mean(xs) - x) / (sigma / Math.sqrt(xs.length)), 0, 1);
});
alias('ZTEST', 'Z.TEST');
fn('PROB', S, 3, 4, 'rrvv', "x_range, prob_range, lower_limit, [upper_limit]", 'Returns the probability that values in a range are between two limits.', (args: Value[], ctx: EvalContext) => {
  const [xs, ps] = pairedNums(args[0], args[1], ctx);
  const total = ps.reduce((a, b) => a + b, 0);
  check(Math.abs(total - 1) < 1e-9 && ps.every((p) => p >= 0 && p <= 1));
  const lo = num(args[2] as Scalar);
  const hi = args.length > 3 && args[3] !== null ? num(args[3] as Scalar) : lo;
  let s = 0;
  xs.forEach((x, i) => {
    if (x >= lo && x <= hi) s += ps[i];
  });
  return s;
});
fn('LOGEST', S, 1, 4, 'r', "known_y's, [known_x's], [const], [stats]", 'Returns the parameters of an exponential trend (m and b of y = b*m^x).', (args: Value[], ctx: EvalContext) => {
  const ys = collectNums([args[0]], ctx);
  const xs = args.length > 1 && args[1] !== null ? collectNums([args[1]], ctx) : ys.map((_, i) => i + 1);
  if (xs.length !== ys.length) throw ERR.REF;
  check(ys.every((y) => y > 0));
  const ly = ys.map(Math.log);
  const mx = mean(xs);
  const my = mean(ly);
  let sxy = 0;
  let sxx = 0;
  xs.forEach((x, i) => {
    sxy += (x - mx) * (ly[i] - my);
    sxx += (x - mx) ** 2;
  });
  if (sxx === 0) throw ERR.DIV0;
  const slope = sxy / sxx;
  return { kind: 'matrix', data: [[Math.exp(slope), Math.exp(my - slope * mx)]] };
});

