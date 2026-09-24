import { ERR, isErr, asErr, type Scalar, type Value } from '../values';
import type { EvalContext } from '../evaluator';
import { parseInput, partsToSerial, serialToParts } from '../../format/numfmt';
import { check, eachArg, fn, int, num, optInt, str } from './helpers';

const D = 'Date & time' as const;

/* ------------------------------------------------------------ helpers */

/** Date serial argument (numbers or date text), must be non-negative. */
function serial(v: Scalar): number {
  const n = num(v);
  if (n < 0 || n >= 2958466) throw ERR.NUM;
  return n;
}

function day(v: Scalar): number {
  return Math.floor(serial(v));
}

function ymd(n: number) {
  return serialToParts(n);
}

const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const daysInMonth = (y: number, m: number) => [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];

function dateSerial(y: number, m: number, d: number): number {
  if (y >= 0 && y < 1900) y += 1900;
  if (y < 0 || y >= 10000) throw ERR.NUM;
  const s = partsToSerial(y, m, d);
  if (s < 0) throw ERR.NUM;
  return s;
}

/* -------------------------------------------------------- constructors */

fn('DATE', D, 3, 3, 'v', 'year, month, day', 'Returns the number that represents a particular date.', ([y, m, d]) => dateSerial(int(y), int(m), int(d)));

fn('TIME', D, 3, 3, 'v', 'hour, minute, second', 'Returns the decimal number for a particular time.', ([h, m, s]) => {
  const total = int(h) * 3600 + int(m) * 60 + int(s);
  if (total < 0) throw ERR.NUM;
  return (total % 86400) / 86400;
});

fn('DATEVALUE', D, 1, 1, 'v', 'date_text', 'Converts a date in the form of text to a serial number.', ([t]) => {
  const p = parseInput(str(t));
  if (typeof p.value !== 'number' || !p.format || !/[dmy]/i.test(p.format)) throw ERR.VALUE;
  return Math.floor(p.value);
});

fn('TIMEVALUE', D, 1, 1, 'v', 'time_text', 'Converts a time in the form of text to a serial number.', ([t]) => {
  const p = parseInput(str(t));
  if (typeof p.value !== 'number' || !p.format || !/[hs]/i.test(p.format)) throw ERR.VALUE;
  return p.value - Math.floor(p.value);
});

fn('NOW', D, 0, 0, 'v', '', 'Returns the current date and time.', (_a, ctx: EvalContext) => ctx.now(), { volatile: true });
fn('TODAY', D, 0, 0, 'v', '', 'Returns the current date.', (_a, ctx: EvalContext) => Math.floor(ctx.now()), { volatile: true });

/* --------------------------------------------------------- components */

fn('YEAR', D, 1, 1, 'v', 'serial_number', 'Returns the year of a date.', ([v]) => ymd(serial(v)).y);
fn('MONTH', D, 1, 1, 'v', 'serial_number', 'Returns the month of a date (1-12).', ([v]) => ymd(serial(v)).m);
fn('DAY', D, 1, 1, 'v', 'serial_number', 'Returns the day of the month (1-31).', ([v]) => ymd(serial(v)).d);
fn('HOUR', D, 1, 1, 'v', 'serial_number', 'Returns the hour (0-23).', ([v]) => ymd(serial(v)).H);
fn('MINUTE', D, 1, 1, 'v', 'serial_number', 'Returns the minute (0-59).', ([v]) => ymd(serial(v)).M);
fn('SECOND', D, 1, 1, 'v', 'serial_number', 'Returns the second (0-59).', ([v]) => ymd(serial(v)).S);

fn('WEEKDAY', D, 1, 2, 'v', 'serial_number, [return_type]', 'Returns a number from 1 to 7 identifying the day of the week.', (args) => {
  const dow = ymd(serial(args[0])).dow; // 0 = Sunday
  const t = optInt(args, 1, 1);
  switch (t) {
    case 1:
    case 17:
      return dow + 1;
    case 2:
    case 11:
      return ((dow + 6) % 7) + 1;
    case 3:
      return (dow + 6) % 7;
    case 12:
    case 13:
    case 14:
    case 15:
    case 16: {
      const start = t - 10; // 12 → Tuesday (2)
      return ((dow - start + 7) % 7) + 1;
    }
  }
  throw ERR.NUM;
});

function isoWeek(n: number): number {
  const p = ymd(n);
  const t = Date.UTC(p.y, p.m - 1, p.d);
  const dt = new Date(t);
  const dayNum = (dt.getUTCDay() + 6) % 7; // Monday = 0
  dt.setUTCDate(dt.getUTCDate() - dayNum + 3);
  const firstThursday = new Date(Date.UTC(dt.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((dt.getTime() - firstThursday.getTime()) / 86400000 - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
}

fn('ISOWEEKNUM', D, 1, 1, 'v', 'date', 'Returns the ISO week number of the year for a given date.', ([v]) => isoWeek(day(v)));

fn('WEEKNUM', D, 1, 2, 'v', 'serial_number, [return_type]', 'Returns the week number in the year.', (args) => {
  const n = day(args[0]);
  const t = optInt(args, 1, 1);
  if (t === 21) return isoWeek(n);
  const startMap: Record<number, number> = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 };
  const start = startMap[t];
  if (start === undefined) throw ERR.NUM;
  const p = ymd(n);
  const jan1 = dateSerial(p.y, 1, 1);
  const jan1dow = ymd(jan1).dow;
  const offset = (jan1dow - start + 7) % 7;
  return Math.floor((n - jan1 + offset) / 7) + 1;
});

/* --------------------------------------------------------- arithmetic */

function addMonths(n: number, months: number): { y: number; m: number; d: number } {
  const p = ymd(n);
  const total = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(total / 12);
  const m = (total % 12) + 1;
  return { y, m, d: Math.min(p.d, daysInMonth(y, m)) };
}

fn('EDATE', D, 2, 2, 'v', 'start_date, months', 'Returns the serial number of the date that is the indicated number of months before or after the start date.', ([s, m]) => {
  const r = addMonths(day(s), int(m));
  return dateSerial(r.y, r.m, r.d);
});

fn('EOMONTH', D, 2, 2, 'v', 'start_date, months', 'Returns the serial number of the last day of the month before or after a specified number of months.', ([s, m]) => {
  const r = addMonths(day(s), int(m));
  return dateSerial(r.y, r.m, daysInMonth(r.y, r.m));
});

fn('DAYS', D, 2, 2, 'v', 'end_date, start_date', 'Returns the number of days between two dates.', ([e, s]) => day(e) - day(s));

fn('DAYS360', D, 2, 3, 'v', 'start_date, end_date, [method]', 'Returns the number of days between two dates based on a 360-day year.', (args) => {
  const a = ymd(day(args[0]));
  const b = ymd(day(args[1]));
  const european = args.length > 2 && args[2] !== null && num(args[2]) !== 0;
  let d1 = a.d;
  let d2 = b.d;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    const lastFebA = a.m === 2 && a.d === daysInMonth(a.y, 2);
    if (d1 === 31 || lastFebA) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
  }
  return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1);
});

fn('DATEDIF', D, 3, 3, 'v', 'start_date, end_date, unit', 'Calculates the number of days, months or years between two dates (units "Y", "M", "D", "MD", "YM", "YD").', ([s, e, u]) => {
  const a = day(s);
  const b = day(e);
  if (a > b) throw ERR.NUM;
  const p = ymd(a);
  const q = ymd(b);
  const unit = str(u).toUpperCase();
  let months = (q.y - p.y) * 12 + (q.m - p.m);
  if (q.d < p.d) months--;
  switch (unit) {
    case 'Y':
      return Math.floor(months / 12);
    case 'M':
      return months;
    case 'D':
      return b - a;
    case 'MD': {
      let d = q.d - p.d;
      if (d < 0) {
        const pm = q.m === 1 ? 12 : q.m - 1;
        const py = q.m === 1 ? q.y - 1 : q.y;
        d += daysInMonth(py, pm);
      }
      return d;
    }
    case 'YM':
      return ((months % 12) + 12) % 12;
    case 'YD': {
      let start = dateSerial(q.y, p.m, p.d);
      if (start > b) start = dateSerial(q.y - 1, p.m, p.d);
      return b - start;
    }
  }
  throw ERR.NUM;
});

/* ------------------------------------------------------- working days */

function weekendMask(code: Scalar | undefined): boolean[] {
  // index 0 = Monday … 6 = Sunday
  if (code === undefined || code === null) return [false, false, false, false, false, true, true];
  if (typeof code === 'string') {
    if (!/^[01]{7}$/.test(code) || code === '1111111') throw ERR.VALUE;
    return [...code].map((c) => c === '1');
  }
  const n = int(code);
  const pairs: Record<number, [number, number]> = { 1: [5, 6], 2: [6, 0], 3: [0, 1], 4: [1, 2], 5: [2, 3], 6: [3, 4], 7: [4, 5] };
  const mask = [false, false, false, false, false, false, false];
  if (pairs[n]) {
    mask[pairs[n][0]] = true;
    mask[pairs[n][1]] = true;
    return mask;
  }
  if (n >= 11 && n <= 17) {
    mask[n - 11] = true;
    return mask;
  }
  throw ERR.NUM;
}

function holidaySet(v: Value | undefined, ctx: EvalContext): Set<number> {
  const out = new Set<number>();
  if (v === undefined || v === null) return out;
  eachArg([v], ctx, (x) => {
    if (x === null) return;
    if (isErr(x)) throw asErr(x);
    out.add(Math.floor(num(x)));
  });
  return out;
}

const isWork = (n: number, mask: boolean[], hol: Set<number>) => !mask[(ymd(n).dow + 6) % 7] && !hol.has(n);

function networkDays(start: number, end: number, mask: boolean[], hol: Set<number>): number {
  const sign = end >= start ? 1 : -1;
  const [a, b] = sign > 0 ? [start, end] : [end, start];
  let count = 0;
  // whole weeks fast path
  const weeks = Math.floor((b - a + 1) / 7);
  const perWeek = mask.filter((x) => !x).length;
  count += weeks * perWeek;
  for (let n = a + weeks * 7; n <= b; n++) if (!mask[(ymd(n).dow + 6) % 7]) count++;
  for (const h of hol) if (h >= a && h <= b && !mask[(ymd(h).dow + 6) % 7]) count--;
  return sign * count;
}

fn('NETWORKDAYS', D, 2, 3, 'vvr', 'start_date, end_date, [holidays]', 'Returns the number of whole working days between two dates.', (args: Value[], ctx) =>
  networkDays(day(args[0] as Scalar), day(args[1] as Scalar), weekendMask(undefined), holidaySet(args[2], ctx)),
);
fn('NETWORKDAYS.INTL', D, 2, 4, 'vvvr', 'start_date, end_date, [weekend], [holidays]', 'Returns the number of whole working days between two dates with custom weekend parameters.', (args: Value[], ctx) =>
  networkDays(day(args[0] as Scalar), day(args[1] as Scalar), weekendMask(args[2] as Scalar | undefined), holidaySet(args[3], ctx)),
);

function workday(start: number, days: number, mask: boolean[], hol: Set<number>): number {
  check(mask.some((x) => !x), ERR.VALUE);
  let n = start;
  const step = days >= 0 ? 1 : -1;
  let left = Math.abs(days);
  while (left > 0) {
    n += step;
    if (isWork(n, mask, hol)) left--;
    if (n < 0 || n > 2958465) throw ERR.NUM;
  }
  return n;
}

fn('WORKDAY', D, 2, 3, 'vvr', 'start_date, days, [holidays]', 'Returns the serial number of the date before or after a specified number of workdays.', (args: Value[], ctx) =>
  workday(day(args[0] as Scalar), int(args[1] as Scalar), weekendMask(undefined), holidaySet(args[2], ctx)),
);
fn('WORKDAY.INTL', D, 2, 4, 'vvvr', 'start_date, days, [weekend], [holidays]', 'Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters.', (args: Value[], ctx) =>
  workday(day(args[0] as Scalar), int(args[1] as Scalar), weekendMask(args[2] as Scalar | undefined), holidaySet(args[3], ctx)),
);

/* ----------------------------------------------------------- YEARFRAC */

export function yearFrac(start: number, end: number, basis: number): number {
  if (start > end) [start, end] = [end, start];
  const a = ymd(start);
  const b = ymd(end);
  switch (basis) {
    case 0: {
      let d1 = a.d;
      let d2 = b.d;
      const lastFebA = a.m === 2 && a.d === daysInMonth(a.y, 2);
      const lastFebB = b.m === 2 && b.d === daysInMonth(b.y, 2);
      if (lastFebA && lastFebB) d2 = 30;
      if (lastFebA) d1 = 30;
      if (d2 === 31 && d1 >= 30) d2 = 30;
      if (d1 === 31) d1 = 30;
      return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360;
    }
    case 1: {
      const days = end - start;
      if (a.y === b.y) return days / (isLeap(a.y) ? 366 : 365);
      const oneYearLater = a.y + 1 < b.y || (a.y + 1 === b.y && (b.m > a.m || (b.m === a.m && b.d > a.d)));
      if (!oneYearLater) {
        const leapInside = (isLeap(a.y) && (a.m < 3)) || (isLeap(b.y) && (b.m > 2 || (b.m === 2 && b.d === 29)));
        return days / (leapInside ? 366 : 365);
      }
      let total = 0;
      for (let y = a.y; y <= b.y; y++) total += isLeap(y) ? 366 : 365;
      return days / (total / (b.y - a.y + 1));
    }
    case 2:
      return (end - start) / 360;
    case 3:
      return (end - start) / 365;
    case 4: {
      const d1 = Math.min(a.d, 30);
      const d2 = Math.min(b.d, 30);
      return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360;
    }
  }
  throw ERR.NUM;
}

fn('YEARFRAC', D, 2, 3, 'v', 'start_date, end_date, [basis]', 'Returns the year fraction representing the number of whole days between two dates.', (args) => yearFrac(day(args[0]), day(args[1]), optInt(args, 2, 0)));
