import { ERR, isErr, asErr, type Scalar, type Value } from '../values';
import type { EvalContext } from '../evaluator';
import { serialToParts, partsToSerial } from '../../format/numfmt';
import { bool, check, eachArg, finite, flat, fn, int, num, optInt, optNum } from './helpers';
import { yearFrac } from './datetime';

const F = 'Financial' as const;

/* --------------------------------------------------- time value of money */

export function fv(rate: number, nper: number, pmt: number, pv: number, type: number): number {
  if (rate === 0) return -(pv + pmt * nper);
  const f = Math.pow(1 + rate, nper);
  return -(pv * f + (pmt * (1 + rate * type) * (f - 1)) / rate);
}

export function pv(rate: number, nper: number, pmt: number, fvv: number, type: number): number {
  if (rate === 0) return -(fvv + pmt * nper);
  const f = Math.pow(1 + rate, nper);
  return -(fvv + (pmt * (1 + rate * type) * (f - 1)) / rate) / f;
}

export function pmt(rate: number, nper: number, pvv: number, fvv: number, type: number): number {
  if (nper === 0) throw ERR.NUM;
  if (rate === 0) return -(pvv + fvv) / nper;
  const f = Math.pow(1 + rate, nper);
  return (-rate * (fvv + pvv * f)) / ((1 + rate * type) * (f - 1));
}

function ipmt(rate: number, per: number, nper: number, pvv: number, fvv: number, type: number): number {
  check(per >= 1 && per <= nper);
  const p = pmt(rate, nper, pvv, fvv, type);
  let interest: number;
  if (per === 1) interest = type === 1 ? 0 : -pvv;
  else interest = type === 1 ? fv(rate, per - 2, p, pvv, 1) - p : fv(rate, per - 1, p, pvv, 0);
  return interest * rate;
}

const typeArg = (args: Scalar[], i: number) => (optNum(args, i, 0) !== 0 ? 1 : 0);

fn('FV', F, 3, 5, 'v', 'rate, nper, pmt, [pv], [type]', 'Returns the future value of an investment based on periodic, constant payments and a constant interest rate.', (a) =>
  finite(fv(num(a[0]), num(a[1]), num(a[2]), optNum(a, 3, 0), typeArg(a, 4))),
);
fn('PV', F, 3, 5, 'v', 'rate, nper, pmt, [fv], [type]', 'Returns the present value of an investment: the total amount that a series of future payments is worth now.', (a) =>
  finite(pv(num(a[0]), num(a[1]), num(a[2]), optNum(a, 3, 0), typeArg(a, 4))),
);
fn('PMT', F, 3, 5, 'v', 'rate, nper, pv, [fv], [type]', 'Calculates the payment for a loan based on constant payments and a constant interest rate.', (a) =>
  finite(pmt(num(a[0]), num(a[1]), num(a[2]), optNum(a, 3, 0), typeArg(a, 4))),
);
fn('IPMT', F, 4, 6, 'v', 'rate, per, nper, pv, [fv], [type]', 'Returns the interest payment for a given period of an investment.', (a) =>
  finite(ipmt(num(a[0]), num(a[1]), num(a[2]), num(a[3]), optNum(a, 4, 0), typeArg(a, 5))),
);
fn('PPMT', F, 4, 6, 'v', 'rate, per, nper, pv, [fv], [type]', 'Returns the payment on the principal for a given period of an investment.', (a) => {
  const rate = num(a[0]);
  const per = num(a[1]);
  const nper = num(a[2]);
  const p = num(a[3]);
  const f = optNum(a, 4, 0);
  const t = typeArg(a, 5);
  return finite(pmt(rate, nper, p, f, t) - ipmt(rate, per, nper, p, f, t));
});

fn('NPER', F, 3, 5, 'v', 'rate, pmt, pv, [fv], [type]', 'Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.', (a) => {
  const rate = num(a[0]);
  const p = num(a[1]);
  const pvv = num(a[2]);
  const fvv = optNum(a, 3, 0);
  const t = typeArg(a, 4);
  if (rate === 0) {
    if (p === 0) throw ERR.NUM;
    return -(pvv + fvv) / p;
  }
  const num1 = p * (1 + rate * t) - fvv * rate;
  const den = p * (1 + rate * t) + pvv * rate;
  if (num1 / den <= 0) throw ERR.NUM;
  return finite(Math.log(num1 / den) / Math.log(1 + rate));
});

fn('RATE', F, 3, 6, 'v', 'nper, pmt, pv, [fv], [type], [guess]', 'Returns the interest rate per period of a loan or an investment.', (a) => {
  const n = num(a[0]);
  const p = num(a[1]);
  const pvv = num(a[2]);
  const fvv = optNum(a, 3, 0);
  const t = typeArg(a, 4);
  let r = optNum(a, 5, 0.1);
  const f = (x: number) => {
    if (Math.abs(x) < 1e-12) return pvv + p * n + fvv;
    const g = Math.pow(1 + x, n);
    return pvv * g + (p * (1 + x * t) * (g - 1)) / x + fvv;
  };
  for (let i = 0; i < 100; i++) {
    const y = f(r);
    const h = Math.max(1e-7, Math.abs(r) * 1e-7);
    const d = (f(r + h) - f(r - h)) / (2 * h);
    if (d === 0 || !Number.isFinite(d)) break;
    const next = r - y / d;
    if (!Number.isFinite(next) || next <= -1) break;
    if (Math.abs(next - r) < 1e-10) return next;
    r = next;
  }
  if (Math.abs(f(r)) < 1e-7) return r;
  throw ERR.NUM;
});

fn('NPV', F, 2, 255, 'vr', 'rate, value1, [value2], …', 'Returns the net present value of an investment based on a discount rate and future cash flows.', (args: Value[], ctx) => {
  const rate = num(args[0] as Scalar);
  let s = 0;
  let i = 1;
  eachArg(args.slice(1), ctx, (v, direct) => {
    if (isErr(v)) throw asErr(v);
    if (typeof v === 'number' || (direct && v !== null)) {
      s += num(v) / Math.pow(1 + rate, i);
      i++;
    }
  });
  return finite(s);
});

function numbersOf(v: Value, ctx: EvalContext): number[] {
  return flat(v, ctx)
    .filter((x) => {
      if (isErr(x)) throw asErr(x);
      return typeof x === 'number';
    })
    .map(Number);
}

function solve(f: (r: number) => number, guess: number): number {
  let r = guess;
  for (let i = 0; i < 200; i++) {
    const y = f(r);
    if (Math.abs(y) < 1e-10) return r;
    const h = 1e-7;
    const d = (f(r + h) - f(r - h)) / (2 * h);
    if (!d || !Number.isFinite(d)) break;
    let next = r - y / d;
    if (next <= -1) next = (r - 1) / 2;
    if (Math.abs(next - r) < 1e-12) return next;
    r = next;
  }
  // fall back to bisection over a wide bracket
  let lo = -0.9999;
  let hi = 10;
  let flo = f(lo);
  if (flo * f(hi) > 0) throw ERR.NUM;
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2;
    const fm = f(mid);
    if (Math.abs(fm) < 1e-10) return mid;
    if (flo * fm < 0) hi = mid;
    else {
      lo = mid;
      flo = fm;
    }
  }
  return (lo + hi) / 2;
}

fn('IRR', F, 1, 2, 'rv', 'values, [guess]', 'Returns the internal rate of return for a series of cash flows.', (args: Value[], ctx) => {
  const vals = numbersOf(args[0], ctx);
  check(vals.some((v) => v > 0) && vals.some((v) => v < 0));
  const guess = args.length > 1 && args[1] !== null ? num(args[1] as Scalar) : 0.1;
  return solve((r) => vals.reduce((s, v, i) => s + v / Math.pow(1 + r, i), 0), guess);
});

function xnpv(rate: number, vals: number[], dates: number[]): number {
  const d0 = dates[0];
  let s = 0;
  for (let i = 0; i < vals.length; i++) s += vals[i] / Math.pow(1 + rate, (dates[i] - d0) / 365);
  return s;
}

fn('XNPV', F, 3, 3, 'vrr', 'rate, values, dates', 'Returns the net present value for a schedule of cash flows that is not necessarily periodic.', (args: Value[], ctx) => {
  const vals = flat(args[1], ctx).map((v) => num(v));
  const dates = flat(args[2], ctx).map((v) => Math.floor(num(v)));
  if (vals.length !== dates.length) throw ERR.NUM;
  return finite(xnpv(num(args[0] as Scalar), vals, dates));
});

fn('XIRR', F, 2, 3, 'rrv', 'values, dates, [guess]', 'Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic.', (args: Value[], ctx) => {
  const vals = flat(args[0], ctx).map((v) => num(v));
  const dates = flat(args[1], ctx).map((v) => Math.floor(num(v)));
  if (vals.length !== dates.length) throw ERR.NUM;
  check(vals.some((v) => v > 0) && vals.some((v) => v < 0));
  const guess = args.length > 2 && args[2] !== null ? num(args[2] as Scalar) : 0.1;
  return solve((r) => xnpv(r, vals, dates), guess);
});

fn('MIRR', F, 3, 3, 'rvv', 'values, finance_rate, reinvest_rate', 'Returns the internal rate of return where positive and negative cash flows are financed at different rates.', (args: Value[], ctx) => {
  const vals = numbersOf(args[0], ctx);
  const fr = num(args[1] as Scalar);
  const rr = num(args[2] as Scalar);
  const n = vals.length;
  let npvPos = 0;
  let npvNeg = 0;
  vals.forEach((v, i) => {
    if (v > 0) npvPos += v / Math.pow(1 + rr, i);
    else npvNeg += v / Math.pow(1 + fr, i);
  });
  if (npvNeg === 0 || npvPos === 0) throw ERR.DIV0;
  return Math.pow((-npvPos * Math.pow(1 + rr, n)) / (npvNeg * (1 + rr)), 1 / (n - 1)) - 1;
});

/* ------------------------------------------------------- depreciation */

fn('SLN', F, 3, 3, 'v', 'cost, salvage, life', 'Returns the straight-line depreciation of an asset for one period.', ([c, s, l]) => {
  const life = num(l);
  if (life === 0) throw ERR.DIV0;
  return (num(c) - num(s)) / life;
});

fn('SYD', F, 4, 4, 'v', 'cost, salvage, life, per', "Returns the sum-of-years' digits depreciation of an asset for a specified period.", ([c, s, l, p]) => {
  const life = num(l);
  const per = num(p);
  check(life > 0 && per > 0 && per <= life);
  return ((num(c) - num(s)) * (life - per + 1) * 2) / (life * (life + 1));
});

fn('DB', F, 4, 5, 'v', 'cost, salvage, life, period, [month]', 'Returns the depreciation of an asset for a specified period using the fixed-declining balance method.', (a) => {
  const cost = num(a[0]);
  const salvage = num(a[1]);
  const life = num(a[2]);
  const period = int(a[3]);
  const month = optInt(a, 4, 12);
  check(cost >= 0 && salvage >= 0 && life > 0 && period > 0 && month >= 1 && month <= 12 && period <= life + (month < 12 ? 1 : 0));
  if (cost === 0) return 0;
  const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000;
  let total = 0;
  let dep = 0;
  for (let p = 1; p <= period; p++) {
    if (p === 1) dep = (cost * rate * month) / 12;
    else if (p === Math.floor(life) + 1) dep = ((cost - total) * rate * (12 - month)) / 12;
    else dep = (cost - total) * rate;
    total += dep;
  }
  return dep;
});

function ddbPeriod(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life;
  let oldValue: number;
  if (rate >= 1) {
    rate = 1;
    oldValue = period === 1 ? cost : 0;
  } else oldValue = cost * Math.pow(1 - rate, period - 1);
  const newValue = cost * Math.pow(1 - rate, period);
  const dep = newValue < salvage ? oldValue - salvage : oldValue - newValue;
  return dep < 0 ? 0 : dep;
}

fn('DDB', F, 4, 5, 'v', 'cost, salvage, life, period, [factor]', 'Returns the depreciation of an asset for a specified period using the double-declining balance method.', (a) => {
  const cost = num(a[0]);
  const salvage = num(a[1]);
  const life = num(a[2]);
  const period = num(a[3]);
  const factor = optNum(a, 4, 2);
  check(cost >= 0 && salvage >= 0 && life > 0 && period > 0 && period <= life && factor > 0);
  return ddbPeriod(cost, salvage, life, period, factor);
});

function interVdb(cost: number, salvage: number, life: number, life1: number, period: number, factor: number): number {
  let vdb = 0;
  const intEnd = Math.ceil(period);
  let salvageValue = cost - salvage;
  let nowSln = false;
  let sln = 0;
  for (let i = 1; i <= intEnd; i++) {
    let term: number;
    if (!nowSln) {
      const ddb = ddbPeriod(cost, salvage, life, i, factor);
      sln = salvageValue / (life1 - (i - 1));
      if (sln > ddb) {
        term = sln;
        nowSln = true;
      } else {
        term = ddb;
        salvageValue -= ddb;
      }
    } else term = sln;
    if (i === intEnd) term *= period + 1 - intEnd;
    vdb += term;
  }
  return vdb;
}

fn('VDB', F, 5, 7, 'v', 'cost, salvage, life, start_period, end_period, [factor], [no_switch]', 'Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method.', (a) => {
  let cost = num(a[0]);
  const salvage = num(a[1]);
  const life = num(a[2]);
  let start = num(a[3]);
  let end = num(a[4]);
  const factor = optNum(a, 5, 2);
  const noSwitch = a.length > 6 && a[6] !== null ? bool(a[6]) : false;
  check(start >= 0 && end >= start && end <= life && cost >= 0 && salvage <= cost && factor > 0);
  if (noSwitch) {
    const intStart = Math.floor(start);
    const intEnd = Math.ceil(end);
    let vdb = 0;
    for (let i = intStart + 1; i <= intEnd; i++) {
      let term = ddbPeriod(cost, salvage, life, i, factor);
      if (i === intStart + 1) term *= Math.min(end, intStart + 1) - start;
      else if (i === intEnd) term *= end + 1 - intEnd;
      vdb += term;
    }
    return vdb;
  }
  if (start !== Math.floor(start) && factor > 1 && start >= life / 2) {
    const part = start - life / 2;
    start = life / 2;
    end -= part;
  }
  cost -= interVdb(cost, salvage, life, life, start, factor);
  return interVdb(cost, salvage, life, life - start, end - start, factor);
});

/* ------------------------------------------------------------- rates */

fn('EFFECT', F, 2, 2, 'v', 'nominal_rate, npery', 'Returns the effective annual interest rate.', ([r, n]) => {
  const rate = num(r);
  const p = Math.trunc(num(n));
  check(rate > 0 && p >= 1);
  return Math.pow(1 + rate / p, p) - 1;
});
fn('NOMINAL', F, 2, 2, 'v', 'effect_rate, npery', 'Returns the annual nominal interest rate.', ([r, n]) => {
  const rate = num(r);
  const p = Math.trunc(num(n));
  check(rate > 0 && p >= 1);
  return p * (Math.pow(1 + rate, 1 / p) - 1);
});
fn('PDURATION', F, 3, 3, 'v', 'rate, pv, fv', 'Returns the number of periods required by an investment to reach a specified value.', ([r, p, f]) => {
  const rate = num(r);
  const a = num(p);
  const b = num(f);
  check(rate > 0 && a > 0 && b > 0);
  return (Math.log(b) - Math.log(a)) / Math.log(1 + rate);
});
fn('RRI', F, 3, 3, 'v', 'nper, pv, fv', 'Returns an equivalent interest rate for the growth of an investment.', ([n, p, f]) => {
  const nper = num(n);
  const a = num(p);
  check(nper > 0 && a !== 0);
  return Math.pow(num(f) / a, 1 / nper) - 1;
});
fn('FVSCHEDULE', F, 2, 2, 'vr', 'principal, schedule', 'Returns the future value of an initial principal after applying a series of compound interest rates.', (args: Value[], ctx) => {
  let v = num(args[0] as Scalar);
  for (const r of flat(args[1], ctx)) v *= 1 + num(r);
  return v;
});
fn('ISPMT', F, 4, 4, 'v', 'rate, per, nper, pv', 'Returns the interest paid during a specific period of an investment (even principal payments).', ([r, p, n, v]) => {
  const nper = num(n);
  if (nper === 0) throw ERR.DIV0;
  return num(v) * num(r) * (num(p) / nper - 1);
});

function cumulative(a: Scalar[], principal: boolean): number {
  const rate = num(a[0]);
  const nper = Math.trunc(num(a[1]));
  const pvv = num(a[2]);
  const start = Math.trunc(num(a[3]));
  const end = Math.trunc(num(a[4]));
  const type = Math.trunc(num(a[5]));
  check(rate > 0 && nper > 0 && pvv > 0 && start >= 1 && end >= start && end <= nper && (type === 0 || type === 1));
  const payment = pmt(rate, nper, pvv, 0, type);
  let s = 0;
  for (let per = start; per <= end; per++) {
    const i = ipmt(rate, per, nper, pvv, 0, type);
    s += principal ? payment - i : i;
  }
  return s;
}

fn('CUMIPMT', F, 6, 6, 'v', 'rate, nper, pv, start_period, end_period, type', 'Returns the cumulative interest paid between two periods.', (a) => cumulative(a, false));
fn('CUMPRINC', F, 6, 6, 'v', 'rate, nper, pv, start_period, end_period, type', 'Returns the cumulative principal paid on a loan between two periods.', (a) => cumulative(a, true));

fn('DOLLARDE', F, 2, 2, 'v', 'fractional_dollar, fraction', 'Converts a dollar price expressed as a fraction into a decimal number.', ([d, f]) => {
  const x = num(d);
  const frac = Math.trunc(num(f));
  check(frac >= 0);
  if (frac === 0) throw ERR.DIV0;
  const digits = Math.ceil(Math.log10(frac));
  const whole = Math.trunc(x);
  return whole + ((x - whole) * Math.pow(10, digits)) / frac;
});
fn('DOLLARFR', F, 2, 2, 'v', 'decimal_dollar, fraction', 'Converts a dollar price expressed as a decimal number into a fraction.', ([d, f]) => {
  const x = num(d);
  const frac = Math.trunc(num(f));
  check(frac >= 0);
  if (frac === 0) throw ERR.DIV0;
  const digits = Math.ceil(Math.log10(frac));
  const whole = Math.trunc(x);
  return whole + ((x - whole) * frac) / Math.pow(10, digits);
});

/* ---------------------------------------------------------- securities */

const lastDay = (y: number, m: number) => new Date(Date.UTC(y, m, 0)).getUTCDate();

function shiftMonths(serial: number, months: number, eom: boolean): number {
  const p = serialToParts(serial);
  const total = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(total / 12);
  const m = (total % 12) + 1;
  const d = eom ? lastDay(y, m) : Math.min(p.d, lastDay(y, m));
  return partsToSerial(y, m, d);
}

function couponDates(settle: number, mat: number, freq: number): { pcd: number; ncd: number; n: number } {
  check(settle < mat && [1, 2, 4].includes(freq));
  const step = 12 / freq;
  const mp = serialToParts(mat);
  const eom = mp.d === lastDay(mp.y, mp.m);
  let n = 0;
  let date = mat;
  while (date > settle) {
    n++;
    date = shiftMonths(mat, -step * n, eom);
  }
  return { pcd: date, ncd: shiftMonths(mat, -step * (n - 1), eom), n };
}

function days360(a: number, b: number, european: boolean): number {
  const p = serialToParts(a);
  const q = serialToParts(b);
  let d1 = p.d;
  let d2 = q.d;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    if (d1 === 31 || (p.m === 2 && d1 === lastDay(p.y, 2))) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
  }
  return (q.y - p.y) * 360 + (q.m - p.m) * 30 + (d2 - d1);
}

function dayCount(a: number, b: number, basis: number): number {
  if (basis === 0) return days360(a, b, false);
  if (basis === 4) return days360(a, b, true);
  return b - a;
}

function couponDays(pcd: number, ncd: number, freq: number, basis: number): number {
  if (basis === 1) return ncd - pcd;
  if (basis === 3) return 365 / freq;
  return 360 / freq;
}

function secArgs(a: Scalar[], freqIdx: number, basisIdx: number) {
  const settle = Math.floor(num(a[0]));
  const mat = Math.floor(num(a[1]));
  const freq = Math.trunc(num(a[freqIdx]));
  const basis = optInt(a, basisIdx, 0);
  check(basis >= 0 && basis <= 4);
  const { pcd, ncd, n } = couponDates(settle, mat, freq);
  const E = couponDays(pcd, ncd, freq, basis);
  const A = dayCount(pcd, settle, basis);
  const DSC = basis === 0 ? E - A : ncd - settle;
  return { settle, mat, freq, basis, pcd, ncd, n, E, A, DSC };
}

fn('COUPNCD', F, 3, 4, 'v', 'settlement, maturity, frequency, [basis]', 'Returns the next coupon date after the settlement date.', (a) => secArgs(a, 2, 3).ncd);
fn('COUPPCD', F, 3, 4, 'v', 'settlement, maturity, frequency, [basis]', 'Returns the previous coupon date before the settlement date.', (a) => secArgs(a, 2, 3).pcd);
fn('COUPNUM', F, 3, 4, 'v', 'settlement, maturity, frequency, [basis]', 'Returns the number of coupons payable between the settlement date and maturity date.', (a) => secArgs(a, 2, 3).n);
fn('COUPDAYBS', F, 3, 4, 'v', 'settlement, maturity, frequency, [basis]', 'Returns the number of days from the beginning of the coupon period to the settlement date.', (a) => secArgs(a, 2, 3).A);
fn('COUPDAYS', F, 3, 4, 'v', 'settlement, maturity, frequency, [basis]', 'Returns the number of days in the coupon period that contains the settlement date.', (a) => secArgs(a, 2, 3).E);
fn('COUPDAYSNC', F, 3, 4, 'v', 'settlement, maturity, frequency, [basis]', 'Returns the number of days from the settlement date to the next coupon date.', (a) => secArgs(a, 2, 3).DSC);

function bondPrice(s: ReturnType<typeof secArgs>, rate: number, yld: number, redemption: number): number {
  const { freq, n, E, A, DSC } = s;
  const cpn = (100 * rate) / freq;
  if (n === 1) return (redemption + cpn) / (1 + ((DSC / E) * yld) / freq) - (cpn * A) / E;
  let price = redemption / Math.pow(1 + yld / freq, n - 1 + DSC / E);
  for (let k = 1; k <= n; k++) price += cpn / Math.pow(1 + yld / freq, k - 1 + DSC / E);
  return price - (cpn * A) / E;
}

fn('PRICE', F, 6, 7, 'v', 'settlement, maturity, rate, yld, redemption, frequency, [basis]', 'Returns the price per $100 face value of a security that pays periodic interest.', (a) => {
  const s = secArgs(a, 5, 6);
  const rate = num(a[2]);
  const yld = num(a[3]);
  const red = num(a[4]);
  check(rate >= 0 && yld >= 0 && red > 0);
  return bondPrice(s, rate, yld, red);
});

fn('YIELD', F, 6, 7, 'v', 'settlement, maturity, rate, pr, redemption, frequency, [basis]', 'Returns the yield on a security that pays periodic interest.', (a) => {
  const s = secArgs(a, 5, 6);
  const rate = num(a[2]);
  const pr = num(a[3]);
  const red = num(a[4]);
  check(rate >= 0 && pr > 0 && red > 0);
  let lo = -0.99;
  let hi = 10;
  const f = (y: number) => bondPrice(s, rate, y, red) - pr;
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2;
    if (f(mid) > 0) lo = mid;
    else hi = mid;
    if (hi - lo < 1e-12) break;
  }
  return (lo + hi) / 2;
});

function duration(a: Scalar[]): number {
  const s = secArgs(a, 4, 5);
  const coupon = num(a[2]);
  const yld = num(a[3]);
  check(coupon >= 0 && yld >= 0);
  const { freq, n, E, DSC } = s;
  const f = DSC / E;
  let pvSum = 0;
  let tSum = 0;
  for (let k = 1; k <= n; k++) {
    const cf = (100 * coupon) / freq + (k === n ? 100 : 0);
    const t = k - 1 + f;
    const disc = cf / Math.pow(1 + yld / freq, t);
    pvSum += disc;
    tSum += (t / freq) * disc;
  }
  return tSum / pvSum;
}

fn('DURATION', F, 5, 6, 'v', 'settlement, maturity, coupon, yld, frequency, [basis]', 'Returns the annual Macaulay duration of a security with periodic interest payments.', (a) => duration(a));
fn('MDURATION', F, 5, 6, 'v', 'settlement, maturity, coupon, yld, frequency, [basis]', 'Returns the modified Macaulay duration for a security with an assumed par value of $100.', (a) => duration(a) / (1 + num(a[3]) / Math.trunc(num(a[4]))));

function yf(a: Scalar, b: Scalar, basis: number): number {
  const s = Math.floor(num(a));
  const m = Math.floor(num(b));
  check(s < m && basis >= 0 && basis <= 4);
  return yearFrac(s, m, basis);
}

fn('ACCRINT', F, 6, 8, 'v', 'issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method]', 'Returns the accrued interest for a security that pays periodic interest.', (a) => {
  const rate = num(a[3]);
  const par = num(a[4]);
  check(rate > 0 && par > 0);
  return par * rate * yf(a[0], a[2], optInt(a, 6, 0));
});
fn('ACCRINTM', F, 3, 5, 'v', 'issue, settlement, rate, [par], [basis]', 'Returns the accrued interest for a security that pays interest at maturity.', (a) => {
  const rate = num(a[2]);
  const par = optNum(a, 3, 1000);
  check(rate > 0 && par > 0);
  return par * rate * yf(a[0], a[1], optInt(a, 4, 0));
});
fn('DISC', F, 4, 5, 'v', 'settlement, maturity, pr, redemption, [basis]', 'Returns the discount rate for a security.', (a) => {
  const pr = num(a[2]);
  const red = num(a[3]);
  check(pr > 0 && red > 0);
  return (red - pr) / red / yf(a[0], a[1], optInt(a, 4, 0));
});
fn('INTRATE', F, 4, 5, 'v', 'settlement, maturity, investment, redemption, [basis]', 'Returns the interest rate for a fully invested security.', (a) => {
  const inv = num(a[2]);
  const red = num(a[3]);
  check(inv > 0 && red > 0);
  return (red - inv) / inv / yf(a[0], a[1], optInt(a, 4, 0));
});
fn('RECEIVED', F, 4, 5, 'v', 'settlement, maturity, investment, discount, [basis]', 'Returns the amount received at maturity for a fully invested security.', (a) => {
  const inv = num(a[2]);
  const disc = num(a[3]);
  check(inv > 0 && disc > 0);
  const d = 1 - disc * yf(a[0], a[1], optInt(a, 4, 0));
  if (d <= 0) throw ERR.NUM;
  return inv / d;
});
fn('PRICEDISC', F, 4, 5, 'v', 'settlement, maturity, discount, redemption, [basis]', 'Returns the price per $100 face value of a discounted security.', (a) => {
  const disc = num(a[2]);
  const red = num(a[3]);
  check(disc > 0 && red > 0);
  return red * (1 - disc * yf(a[0], a[1], optInt(a, 4, 0)));
});
fn('PRICEMAT', F, 5, 6, 'v', 'settlement, maturity, issue, rate, yld, [basis]', 'Returns the price per $100 face value of a security that pays interest at maturity.', (a) => {
  const basis = optInt(a, 5, 0);
  const rate = num(a[3]);
  const yld = num(a[4]);
  check(rate >= 0 && yld >= 0);
  const dsm = yf(a[0], a[1], basis);
  const dim = yf(a[2], a[1], basis);
  const aa = yf(a[2], a[0], basis);
  return (100 + dim * rate * 100) / (1 + dsm * yld) - aa * rate * 100;
});
fn('YIELDDISC', F, 4, 5, 'v', 'settlement, maturity, pr, redemption, [basis]', 'Returns the annual yield for a discounted security.', (a) => {
  const pr = num(a[2]);
  const red = num(a[3]);
  check(pr > 0 && red > 0);
  return (red - pr) / pr / yf(a[0], a[1], optInt(a, 4, 0));
});
fn('YIELDMAT', F, 5, 6, 'v', 'settlement, maturity, issue, rate, pr, [basis]', 'Returns the annual yield of a security that pays interest at maturity.', (a) => {
  const basis = optInt(a, 5, 0);
  const rate = num(a[3]);
  const pr = num(a[4]);
  check(rate >= 0 && pr > 0);
  const dsm = yf(a[0], a[1], basis);
  const dim = yf(a[2], a[1], basis);
  const aa = yf(a[2], a[0], basis);
  return ((1 + dim * rate - (pr / 100 + aa * rate)) / (pr / 100 + aa * rate)) / dsm;
});
fn('TBILLPRICE', F, 3, 3, 'v', 'settlement, maturity, discount', 'Returns the price per $100 face value for a Treasury bill.', ([s, m, d]) => {
  const dsm = Math.floor(num(m)) - Math.floor(num(s));
  const disc = num(d);
  check(dsm > 0 && dsm <= 366 && disc > 0);
  return 100 * (1 - (disc * dsm) / 360);
});
fn('TBILLYIELD', F, 3, 3, 'v', 'settlement, maturity, pr', 'Returns the yield for a Treasury bill.', ([s, m, p]) => {
  const dsm = Math.floor(num(m)) - Math.floor(num(s));
  const pr = num(p);
  check(dsm > 0 && dsm <= 366 && pr > 0);
  return ((100 - pr) / pr) * (360 / dsm);
});
fn('TBILLEQ', F, 3, 3, 'v', 'settlement, maturity, discount', 'Returns the bond-equivalent yield for a Treasury bill.', ([s, m, d]) => {
  const dsm = Math.floor(num(m)) - Math.floor(num(s));
  const disc = num(d);
  check(dsm > 0 && dsm <= 366 && disc > 0);
  return (365 * disc) / (360 - disc * dsm);
});
