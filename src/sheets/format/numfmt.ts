/**
 * Excel-compatible number formatting.
 * Supports sections (pos;neg;zero;text), [conditions], [colours], literals ("…", \x),
 * 0 # ? placeholders, thousands separators and scaling, %, scientific notation,
 * fractions, dates and times (including [h]:mm elapsed time and AM/PM) and General.
 */

/* ---------------------------------------------------------------- dates */

const EPOCH = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

export interface DateParts {
  y: number;
  m: number; // 1-12
  d: number;
  H: number;
  M: number;
  S: number;
  ms: number;
  dow: number; // 0 = Sunday
}

/** Excel serial (1900 date system, including the Feb-29-1900 quirk) → calendar parts. */
export function serialToParts(serial: number): DateParts {
  let days = Math.floor(serial);
  let frac = serial - days;
  // round to the millisecond to avoid 23:59:59.999 artefacts
  let msTotal = Math.round(frac * DAY_MS);
  if (msTotal >= DAY_MS) {
    days += 1;
    msTotal -= DAY_MS;
  }
  frac = msTotal;
  let y: number;
  let m: number;
  let d: number;
  let dow: number;
  if (days === 60) {
    y = 1900;
    m = 2;
    d = 29;
    dow = 3;
  } else {
    const base = days < 60 ? Date.UTC(1899, 11, 31) : EPOCH;
    const dt = new Date(base + days * DAY_MS);
    y = dt.getUTCFullYear();
    m = dt.getUTCMonth() + 1;
    d = dt.getUTCDate();
    dow = dt.getUTCDay();
  }
  const H = Math.floor(frac / 3_600_000);
  const M = Math.floor((frac % 3_600_000) / 60_000);
  const S = Math.floor((frac % 60_000) / 1000);
  const ms = frac % 1000;
  return { y, m, d, H, M, S, ms, dow };
}

export function partsToSerial(y: number, m: number, d: number, H = 0, M = 0, S = 0): number {
  // Date.UTC handles overflow of months/days like Excel's DATE()
  const t = Date.UTC(y, m - 1, d);
  let days = Math.round((t - EPOCH) / DAY_MS);
  if (days < 61) days -= 1; // Excel's fictional 1900-02-29
  return days + (H * 3600 + M * 60 + S) / 86400;
}

export function dateToSerial(dt: Date): number {
  return partsToSerial(dt.getFullYear(), dt.getMonth() + 1, dt.getDate(), dt.getHours(), dt.getMinutes(), dt.getSeconds() + dt.getMilliseconds() / 1000);
}

export function serialToDate(serial: number): Date {
  const p = serialToParts(serial);
  return new Date(p.y, p.m - 1, p.d, p.H, p.M, p.S, p.ms);
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/* ------------------------------------------------------------- tokens */

export type FormatTok =
  | { t: 'lit'; v: string }
  | { t: 'digit'; v: '0' | '#' | '?' }
  | { t: 'point' }
  | { t: 'comma' }
  | { t: 'percent' }
  | { t: 'exp'; sign: '+' | '-' }
  | { t: 'slash' }
  | { t: 'text' }
  | { t: 'date'; v: string }
  | { t: 'ampm'; v: string }
  | { t: 'elapsed'; v: 'h' | 'm' | 's'; len: number }
  | { t: 'subsec'; len: number }
  | { t: 'general' };

export interface FormatSection {
  toks: FormatTok[];
  color?: string;
  cond?: { op: string; v: number };
  isDate: boolean;
  isText: boolean;
  isGeneral: boolean;
  hasAmPm: boolean;
}

const COLORS: Record<string, string> = {
  black: '#000000',
  white: '#ffffff',
  red: '#ff0000',
  green: '#00a000',
  blue: '#0000ff',
  yellow: '#c0a000',
  magenta: '#ff00ff',
  cyan: '#00a0a0',
};

const INDEXED = ['#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080'];

function splitSections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inQ = false;
  let inB = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '\\' && !inQ) {
      cur += ch + (fmt[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"') inQ = !inQ;
    else if (ch === '[' && !inQ) inB = true;
    else if (ch === ']' && !inQ) inB = false;
    if (ch === ';' && !inQ && !inB) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

function parseSection(src: string): FormatSection {
  const toks: FormatTok[] = [];
  const sec: FormatSection = { toks, isDate: false, isText: false, isGeneral: false, hasAmPm: false };
  let i = 0;
  const lower = src.toLowerCase();
  while (i < src.length) {
    const ch = src[i];
    const lc = lower[i];
    if (ch === '"') {
      const j = src.indexOf('"', i + 1);
      toks.push({ t: 'lit', v: src.slice(i + 1, j < 0 ? undefined : j) });
      i = j < 0 ? src.length : j + 1;
      continue;
    }
    if (ch === '\\') {
      toks.push({ t: 'lit', v: src[i + 1] ?? '' });
      i += 2;
      continue;
    }
    if (ch === '_') {
      toks.push({ t: 'lit', v: ' ' });
      i += 2;
      continue;
    }
    if (ch === '*') {
      i += 2;
      continue;
    }
    if (ch === '[') {
      const j = src.indexOf(']', i);
      const inner = src.slice(i + 1, j < 0 ? undefined : j);
      i = j < 0 ? src.length : j + 1;
      const il = inner.toLowerCase();
      const cm = /^([<>=]=?|<>)\s*(-?[\d.]+)$/.exec(inner);
      if (cm) sec.cond = { op: cm[1], v: parseFloat(cm[2]) };
      else if (COLORS[il]) sec.color = COLORS[il];
      else if (/^color\s*\d+$/.test(il)) sec.color = INDEXED[(parseInt(il.replace(/\D/g, ''), 10) - 1) % INDEXED.length];
      else if (/^h+$/.test(il)) toks.push({ t: 'elapsed', v: 'h', len: il.length });
      else if (/^m+$/.test(il)) toks.push({ t: 'elapsed', v: 'm', len: il.length });
      else if (/^s+$/.test(il)) toks.push({ t: 'elapsed', v: 's', len: il.length });
      else if (inner.startsWith('$')) {
        const sym = inner.slice(1).split('-')[0];
        if (sym) toks.push({ t: 'lit', v: sym });
      }
      if (/^[hms]+$/.test(il)) sec.isDate = true;
      continue;
    }
    if (lower.startsWith('general', i)) {
      toks.push({ t: 'general' });
      sec.isGeneral = true;
      i += 7;
      continue;
    }
    if (lower.startsWith('am/pm', i) || lower.startsWith('a/p', i)) {
      const len = lower.startsWith('am/pm', i) ? 5 : 3;
      toks.push({ t: 'ampm', v: src.slice(i, i + len) });
      sec.hasAmPm = true;
      sec.isDate = true;
      i += len;
      continue;
    }
    if (lc === '0' || ch === '#' || ch === '?') {
      toks.push({ t: 'digit', v: ch as '0' | '#' | '?' });
      i++;
      continue;
    }
    if (ch === '.') {
      // fractional seconds: ".0", ".00" directly after seconds
      const prev = toks[toks.length - 1];
      if (prev && prev.t === 'date' && prev.v.startsWith('s') && lower[i + 1] === '0') {
        let n = 0;
        while (lower[i + 1 + n] === '0') n++;
        toks.push({ t: 'subsec', len: n });
        i += 1 + n;
        continue;
      }
      toks.push({ t: 'point' });
      i++;
      continue;
    }
    if (ch === ',') {
      toks.push({ t: 'comma' });
      i++;
      continue;
    }
    if (ch === '%') {
      toks.push({ t: 'percent' });
      i++;
      continue;
    }
    if ((lc === 'e') && (src[i + 1] === '+' || src[i + 1] === '-') && toks.some((t) => t.t === 'digit')) {
      toks.push({ t: 'exp', sign: src[i + 1] as '+' | '-' });
      i += 2;
      continue;
    }
    if (ch === '/' && toks.some((t) => t.t === 'digit') && /[0#?\d]/.test(src[i + 1] ?? '')) {
      toks.push({ t: 'slash' });
      i++;
      continue;
    }
    if (ch === '@') {
      toks.push({ t: 'text' });
      sec.isText = true;
      i++;
      continue;
    }
    if ('ymdhs'.includes(lc)) {
      let j = i;
      while (j < src.length && lower[j] === lc) j++;
      toks.push({ t: 'date', v: lower.slice(i, j) });
      sec.isDate = true;
      i = j;
      continue;
    }
    if (lc === 'b' && (lower.startsWith('bb', i) || lower.startsWith('bbbb', i))) {
      let j = i;
      while (lower[j] === 'b') j++;
      toks.push({ t: 'date', v: 'y'.repeat(j - i) });
      sec.isDate = true;
      i = j;
      continue;
    }
    if (/\d/.test(ch) && toks.some((t) => t.t === 'slash')) {
      // fixed denominator digits
      let j = i;
      while (/\d/.test(src[j] ?? '')) j++;
      toks.push({ t: 'lit', v: `\u0000${src.slice(i, j)}` });
      i = j;
      continue;
    }
    toks.push({ t: 'lit', v: ch });
    i++;
  }
  // resolve m → minutes when next to h or s
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    if (t.t !== 'date' || !t.v.startsWith('m') || t.v.length > 2) continue;
    let prev: FormatTok | undefined;
    for (let p = k - 1; p >= 0; p--) if (toks[p].t === 'date' || toks[p].t === 'elapsed') {
      prev = toks[p];
      break;
    }
    let next: FormatTok | undefined;
    for (let n = k + 1; n < toks.length; n++) if (toks[n].t === 'date' || toks[n].t === 'elapsed') {
      next = toks[n];
      break;
    }
    const isH = (x?: FormatTok) => (x?.t === 'date' && x.v.startsWith('h')) || (x?.t === 'elapsed' && x.v === 'h');
    const isS = (x?: FormatTok) => (x?.t === 'date' && x.v.startsWith('s')) || (x?.t === 'elapsed' && x.v === 's');
    if (isH(prev) || isS(next)) toks[k] = { t: 'date', v: t.v === 'mm' ? 'MM' : 'M' };
  }
  return sec;
}

const cache = new Map<string, FormatSection[]>();

/** A format code split into sections (positive; negative; zero; text) of tokens. */
export function sectionsFor(fmt: string): FormatSection[] {
  let s = cache.get(fmt);
  if (!s) {
    s = splitSections(fmt).map(parseSection);
    if (cache.size > 2000) cache.clear();
    cache.set(fmt, s);
  }
  return s;
}

/* ------------------------------------------------------------ general */

export function formatGeneral(v: number): string {
  if (!Number.isFinite(v)) return v > 0 ? '#NUM!' : '#NUM!';
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  const abs = Math.abs(v);
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) {
    const [m, e] = v.toExponential(5).split('e');
    const mant = m.replace(/\.?0+$/, '');
    const exp = Number(e);
    return `${mant}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  // up to 10 significant digits, no exponent
  const p = parseFloat(v.toPrecision(10));
  let s = String(p);
  if (s.includes('e')) s = p.toFixed(Math.min(20, Math.max(0, 10 - Math.floor(Math.log10(abs)) - 1))).replace(/\.?0+$/, '');
  return s;
}

/* ---------------------------------------------------------- numbers */

function groupThousands(s: string): string {
  return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function roundTo(v: number, dec: number): number {
  if (dec > 15) return v;
  const f = Math.pow(10, dec);
  return Math.round((v + Number.EPSILON * Math.sign(v)) * f) / f;
}

function gcd(a: number, b: number): number {
  return b ? gcd(b, a % b) : a;
}

function approxFraction(x: number, maxDen: number): [number, number] {
  // best rational approximation with denominator <= maxDen (Stern-Brocot / continued fractions)
  let bestN = Math.round(x);
  let bestD = 1;
  let bestErr = Math.abs(x - bestN);
  for (let d = 1; d <= maxDen; d++) {
    const n = Math.round(x * d);
    const err = Math.abs(x - n / d);
    if (err < bestErr - 1e-12) {
      bestErr = err;
      bestN = n;
      bestD = d;
    }
  }
  const g = gcd(Math.abs(bestN), bestD) || 1;
  return [bestN / g, bestD / g];
}

function formatNumberSection(sec: FormatSection, value: number, negSection: boolean): string {
  const toks = sec.toks;
  let v = value;
  const percents = toks.filter((t) => t.t === 'percent').length;
  v *= Math.pow(100, percents);

  // scaling: commas right after the last digit placeholder divide by 1000 each ("0,," or "0.0,,")
  let lastDigit = -1;
  const pointIdx = toks.findIndex((t) => t.t === 'point');
  const expIdx = toks.findIndex((t) => t.t === 'exp');
  const slashIdx = toks.findIndex((t) => t.t === 'slash');
  const intEnd = pointIdx >= 0 ? pointIdx : expIdx >= 0 ? expIdx : toks.length;
  const numEnd = expIdx >= 0 ? expIdx : toks.length;
  for (let i = 0; i < numEnd; i++) if (toks[i].t === 'digit') lastDigit = i;
  let scale = 0;
  for (let i = lastDigit + 1; i < numEnd && toks[i]?.t === 'comma'; i++) scale++;
  v /= Math.pow(1000, scale);
  let lastIntDigit = -1;
  for (let i = 0; i < intEnd; i++) if (toks[i].t === 'digit') lastIntDigit = i;
  const hasThousands = toks.some((t, i) => t.t === 'comma' && i < lastIntDigit && toks.slice(0, i).some((x) => x.t === 'digit'));

  const neg = v < 0 && !negSection;
  let abs = Math.abs(v);

  /* ---------- fractions */
  if (slashIdx >= 0) {
    // find numerator/denominator placeholder groups
    let numStart = slashIdx - 1;
    while (numStart > 0 && toks[numStart - 1].t === 'digit') numStart--;
    const hasWhole = toks.slice(0, numStart).some((t) => t.t === 'digit');
    const denToks: FormatTok[] = [];
    let k = slashIdx + 1;
    while (k < toks.length && (toks[k].t === 'digit' || (toks[k].t === 'lit' && (toks[k] as { v: string }).v.startsWith('\u0000')))) denToks.push(toks[k++]);
    const fixedDen = denToks.find((t) => t.t === 'lit') as { v: string } | undefined;
    const whole = hasWhole ? Math.floor(abs) : 0;
    let frac = hasWhole ? abs - whole : abs;
    let num: number;
    let den: number;
    if (fixedDen) {
      den = parseInt(fixedDen.v.slice(1), 10) || 1;
      num = Math.round(frac * den);
    } else {
      const maxDen = Math.pow(10, denToks.length) - 1;
      [num, den] = approxFraction(frac, Math.max(1, maxDen));
    }
    let w = whole;
    if (num === den && hasWhole) {
      w += 1;
      num = 0;
    }
    frac = 0;
    const numWidth = slashIdx - numStart;
    const denWidth = denToks.length;
    let out = neg ? '-' : '';
    for (let i = 0; i < numStart; i++) {
      const t = toks[i];
      if (t.t === 'lit') out += t.v;
    }
    if (hasWhole) out += w || num === 0 ? String(w) : '';
    if (hasWhole && out.length && !/\s$/.test(out) && num !== 0) out += ' ';
    if (num === 0 && hasWhole) {
      out += ' '.repeat(numWidth + 1 + denWidth);
    } else {
      out += String(num).padStart(numWidth, ' ') + '/' + (fixedDen ? String(den) : String(den).padEnd(denWidth, ' '));
    }
    for (let i = k; i < toks.length; i++) {
      const t = toks[i];
      if (t.t === 'lit') out += t.v;
    }
    return out;
  }

  /* ---------- scientific */
  let exponent = 0;
  if (expIdx >= 0) {
    const intDigits = toks.slice(0, pointIdx >= 0 ? pointIdx : expIdx).filter((t) => t.t === 'digit').length || 1;
    if (abs !== 0) {
      exponent = Math.floor(Math.log10(abs));
      // engineering-style grouping when several integer placeholders (e.g. ##0.0E+0)
      if (intDigits > 1) exponent = Math.floor(exponent / intDigits) * intDigits;
      else exponent -= intDigits - 1;
      abs = abs / Math.pow(10, exponent);
    }
  }

  const decToks = pointIdx >= 0 ? toks.slice(pointIdx + 1, expIdx >= 0 ? expIdx : toks.length).filter((t) => t.t === 'digit') : [];
  const decimals = decToks.length;
  abs = roundTo(abs, decimals);
  if (expIdx >= 0 && abs >= 10 && exponent !== 0) {
    // rounding overflow (9.99 → 10.0)
    const intDigits = toks.slice(0, pointIdx >= 0 ? pointIdx : expIdx).filter((t) => t.t === 'digit').length || 1;
    if (Math.floor(Math.log10(abs)) + 1 > intDigits) {
      abs /= 10;
      exponent += 1;
    }
  }
  const intPart = Math.floor(abs);
  let intStr = intPart === 0 ? '' : intPart.toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 0 });
  let fracStr = decimals ? (abs - intPart).toFixed(decimals).slice(2) : '';
  if (fracStr.length > decimals) fracStr = fracStr.slice(0, decimals);

  const intToks = toks.slice(0, intEnd).map((t, i) => ({ t, i })).filter((x) => x.t.t === 'digit');
  const contiguous = intToks.length > 0 && intToks.every((x, j) => j === 0 || toks.slice(intToks[j - 1].i + 1, x.i).every((t) => t.t === 'comma'));

  // integer digits
  const intOut = new Map<number, string>();
  if (contiguous) {
    const zeros = intToks.filter((x) => (x.t as { v: string }).v === '0').length;
    const qs = intToks.filter((x) => (x.t as { v: string }).v === '?').length;
    let s = intStr;
    if (s.length < zeros) s = s.padStart(zeros, '0');
    if (hasThousands) s = groupThousands(s);
    if (s.length < zeros + qs) s = s.padStart(zeros + qs, ' ');
    intToks.forEach((x, j) => intOut.set(x.i, j === 0 ? s : ''));
  } else {
    let digits = intStr;
    for (let j = intToks.length - 1; j >= 0; j--) {
      const ph = (intToks[j].t as { v: string }).v;
      let d = '';
      if (digits.length) {
        d = digits[digits.length - 1];
        digits = digits.slice(0, -1);
      } else d = ph === '0' ? '0' : ph === '?' ? ' ' : '';
      if (j === 0 && digits.length) d = digits + d;
      intOut.set(intToks[j].i, d);
    }
  }

  // decimal digits
  const decOut = new Map<number, string>();
  if (pointIdx >= 0) {
    const decIdx = toks.map((t, i) => ({ t, i })).filter((x) => x.i > pointIdx && (expIdx < 0 || x.i < expIdx) && x.t.t === 'digit');
    // trailing zeros may be dropped for # and ?
    let lastSignificant = fracStr.length - 1;
    while (lastSignificant >= 0 && fracStr[lastSignificant] === '0') lastSignificant--;
    decIdx.forEach((x, j) => {
      const ph = (x.t as { v: string }).v;
      const d = fracStr[j] ?? '0';
      if (j > lastSignificant) decOut.set(x.i, ph === '0' ? d : ph === '?' ? ' ' : '');
      else decOut.set(x.i, d);
    });
  }

  let out = '';
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    switch (t.t) {
      case 'lit':
        out += t.v;
        break;
      case 'digit':
        out += intOut.get(i) ?? decOut.get(i) ?? '';
        break;
      case 'point':
        out += '.';
        break;
      case 'comma':
        break;
      case 'percent':
        out += '%';
        break;
      case 'exp': {
        const e = exponent;
        const expDigits = toks.slice(i + 1).filter((x) => x.t === 'digit').length || 1;
        out += `E${e < 0 ? '-' : t.sign === '+' ? '+' : ''}${String(Math.abs(e)).padStart(expDigits, '0')}`;
        // skip exponent digit placeholders
        while (toks[i + 1]?.t === 'digit') i++;
        break;
      }
      case 'text':
        break;
      case 'general':
        out += formatGeneral(abs);
        break;
      default:
        break;
    }
  }
  // Excel puts the minus sign in front of everything (e.g. -$1,234.50)
  return neg ? `-${out}` : out;
}

/* ------------------------------------------------------------ dates */

function formatDateSection(sec: FormatSection, value: number): string {
  const hasSub = sec.toks.find((t) => t.t === 'subsec') as { len: number } | undefined;
  // round to displayed precision
  const precision = hasSub ? Math.pow(10, hasSub.len) : 1;
  const secs = Math.round(value * 86400 * precision) / precision;
  const v = secs / 86400;
  const p = serialToParts(v);
  if (!hasSub) p.ms = 0;
  let out = '';
  for (const t of sec.toks) {
    switch (t.t) {
      case 'lit':
        out += t.v;
        break;
      case 'date': {
        const k = t.v;
        if (k.startsWith('y')) out += k.length <= 2 ? String(p.y % 100).padStart(2, '0') : String(p.y);
        else if (k === 'm') out += String(p.m);
        else if (k === 'mm') out += String(p.m).padStart(2, '0');
        else if (k === 'mmm') out += MONTHS[p.m - 1].slice(0, 3);
        else if (k === 'mmmmm') out += MONTHS[p.m - 1][0];
        else if (k.startsWith('mmmm')) out += MONTHS[p.m - 1];
        else if (k === 'M') out += String(p.M);
        else if (k === 'MM') out += String(p.M).padStart(2, '0');
        else if (k === 'd') out += String(p.d);
        else if (k === 'dd') out += String(p.d).padStart(2, '0');
        else if (k === 'ddd') out += DAYS[p.dow].slice(0, 3);
        else if (k.startsWith('dddd')) out += DAYS[p.dow];
        else if (k.startsWith('h')) {
          let h = p.H;
          if (sec.hasAmPm) h = h % 12 || 12;
          out += k.length >= 2 ? String(h).padStart(2, '0') : String(h);
        } else if (k.startsWith('s')) out += k.length >= 2 ? String(p.S).padStart(2, '0') : String(p.S);
        break;
      }
      case 'ampm': {
        const pm = p.H >= 12;
        if (t.v.length === 5) out += t.v[0] === 'a' || t.v[0] === 'A' ? (pm ? (t.v[0] === 'a' ? 'pm' : 'PM') : t.v[0] === 'a' ? 'am' : 'AM') : pm ? 'PM' : 'AM';
        else out += pm ? (t.v[0] === 'a' ? 'p' : 'P') : t.v[0] === 'a' ? 'a' : 'A';
        break;
      }
      case 'elapsed': {
        const total = Math.abs(v);
        const n = t.v === 'h' ? Math.floor(total * 24) : t.v === 'm' ? Math.floor(total * 1440) : Math.floor(total * 86400);
        out += String(n).padStart(t.len, '0');
        break;
      }
      case 'subsec':
        out += `.${String(Math.floor((p.ms / 1000) * Math.pow(10, t.len))).padStart(t.len, '0')}`;
        break;
      case 'digit':
        out += t.v === '0' ? '0' : '';
        break;
      case 'point':
        out += '.';
        break;
      case 'comma':
        out += ',';
        break;
      case 'percent':
        out += '%';
        break;
      default:
        break;
    }
  }
  return out;
}

/* ------------------------------------------------------------- public */

export interface Formatted {
  text: string;
  color?: string;
}

function condMatch(c: { op: string; v: number }, v: number): boolean {
  switch (c.op) {
    case '<':
      return v < c.v;
    case '<=':
      return v <= c.v;
    case '>':
      return v > c.v;
    case '>=':
      return v >= c.v;
    case '=':
    case '==':
      return v === c.v;
    case '<>':
      return v !== c.v;
    default:
      return false;
  }
}

export function formatValue(value: unknown, fmt: string | undefined | null): Formatted {
  if (value === null || value === undefined || value === '') return { text: '' };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' };
  if (typeof value === 'object' && value && 'error' in (value as Record<string, unknown>)) return { text: String((value as { error: string }).error) };
  const f = fmt && fmt.trim() ? fmt : 'General';
  if (typeof value === 'string') {
    const secs = sectionsFor(f);
    const textSec = secs.find((s) => s.isText) ?? (secs.length >= 4 ? secs[3] : undefined);
    if (!textSec) return { text: value };
    let out = '';
    for (const t of textSec.toks) {
      if (t.t === 'text') out += value;
      else if (t.t === 'lit') out += t.v;
    }
    return { text: out, color: textSec.color };
  }
  if (typeof value !== 'number' || Number.isNaN(value)) return { text: String(value) };
  if (!Number.isFinite(value)) return { text: '#NUM!' };
  if (f === 'General' || f === 'general') return { text: formatGeneral(value) };
  if (f === '@') return { text: formatGeneral(value) };
  const secs = sectionsFor(f).filter((s) => !s.isText || s.toks.some((t) => t.t !== 'text' && t.t !== 'lit'));
  const all = sectionsFor(f);
  let sec: FormatSection;
  let negSection = false;
  const conds = all.filter((s) => s.cond);
  if (conds.length) {
    const hit = all.find((s) => s.cond && condMatch(s.cond, value));
    if (hit) sec = hit;
    else sec = all.find((s) => !s.cond && !s.isText) ?? all[all.length - 1];
    negSection = value < 0 && sec !== all[0];
  } else if (value < 0 && secs.length >= 2) {
    sec = secs[1];
    negSection = true;
  } else if (value === 0 && secs.length >= 3) sec = secs[2];
  else sec = secs[0] ?? all[0];
  if (sec.isGeneral && !sec.toks.some((t) => t.t === 'lit')) return { text: formatGeneral(negSection ? Math.abs(value) : value), color: sec.color };
  if (sec.isDate) {
    if (value < 0) return { text: '#'.repeat(8), color: sec.color };
    return { text: formatDateSection(sec, value), color: sec.color };
  }
  const v = negSection ? Math.abs(value) : value;
  return { text: formatNumberSection(sec, v, negSection), color: sec.color };
}

export function isDateFormat(fmt: string | undefined | null): boolean {
  if (!fmt) return false;
  return sectionsFor(fmt)[0]?.isDate ?? false;
}

export function isPercentFormat(fmt: string | undefined | null): boolean {
  return Boolean(fmt && /%/.test(fmt.replace(/"[^"]*"/g, '')));
}

/* ---------------------------------------------------------- built-ins */

/** Excel's built-in number format ids (used by XLSX files). */
export const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

/* --------------------------------------------------------- input parse */

export interface ParsedInput {
  value: number | string | boolean | null;
  format?: string;
}

const MONTH_NAMES = MONTHS.map((m) => m.toLowerCase());

function monthFromName(s: string): number {
  const l = s.toLowerCase();
  const i = MONTH_NAMES.findIndex((m) => m === l || m.slice(0, 3) === l.slice(0, 3));
  return i + 1;
}

function dayFirstLocale(): boolean {
  try {
    const sample = new Date(2001, 11, 25).toLocaleDateString();
    return sample.startsWith('25');
  } catch {
    return false;
  }
}

const DAY_FIRST = typeof navigator !== 'undefined' ? dayFirstLocale() : false;

/** Short date format matching the user's locale order. */
export const SHORT_DATE = DAY_FIRST ? 'dd/mm/yyyy' : 'm/d/yyyy';

/** Interprets typed cell input like Excel: numbers, %, currency, dates, times, booleans. */
export function parseInput(raw: string): ParsedInput {
  const s = raw.trim();
  if (raw === '') return { value: null };
  if (raw.startsWith("'")) return { value: raw.slice(1) };
  const up = s.toUpperCase();
  if (up === 'TRUE') return { value: true };
  if (up === 'FALSE') return { value: false };

  // numbers with optional currency, thousands, percent, parentheses for negatives
  const num = /^([-+]?)\(?\s*([$€£¥₹]?)\s*([-+]?)((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE]([-+]?\d+))?\s*([$€£¥₹]?)\)?\s*(%?)$/.exec(s);
  if (num) {
    const paren = s.startsWith('(') && s.endsWith(')');
    const sign = num[1] === '-' || num[3] === '-' || paren ? -1 : 1;
    const digits = num[4].replace(/,/g, '');
    let v = parseFloat(digits) * (num[5] ? Math.pow(10, Number(num[5])) : 1) * sign;
    const cur = num[2] || num[6];
    const pct = num[7] === '%';
    if (pct) v /= 100;
    const decs = digits.includes('.') ? digits.split('.')[1].length : 0;
    let format: string | undefined;
    if (pct) format = decs ? `0.${'0'.repeat(Math.min(decs, 10))}%` : '0%';
    else if (cur) format = `${cur === '$' ? '$' : `"${cur}"`}#,##0${decs ? '.00' : ''}`;
    else if (num[4].includes(',')) format = decs ? '#,##0.00' : '#,##0';
    else if (num[5]) format = '0.00E+00';
    return { value: v, format };
  }

  // time only: 14:30, 2:30 pm, 14:30:15
  const tm = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([ap]\.?m?\.?)?$/i.exec(s);
  if (tm) {
    let h = Number(tm[1]);
    const mi = Number(tm[2]);
    const se = tm[3] ? parseFloat(tm[3]) : 0;
    const ap = tm[4]?.toLowerCase();
    if (ap) {
      if (ap.startsWith('p') && h < 12) h += 12;
      if (ap.startsWith('a') && h === 12) h = 0;
    }
    if (h < 24 && mi < 60 && se < 60) {
      return { value: (h * 3600 + mi * 60 + se) / 86400, format: ap ? (tm[3] ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : tm[3] ? 'h:mm:ss' : 'h:mm' };
    }
  }

  // ISO date (+time)
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  if (iso) {
    const [y, m, d] = [Number(iso[1]), Number(iso[2]), Number(iso[3])];
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      const v = partsToSerial(y, m, d, Number(iso[4] ?? 0), Number(iso[5] ?? 0), Number(iso[6] ?? 0));
      return { value: v, format: iso[4] ? 'yyyy-mm-dd h:mm' : 'yyyy-mm-dd' };
    }
  }
  // numeric dates: 3/15/2024, 15/03/2024, 15.03.2024
  const nd = /^(\d{1,2})([/.-])(\d{1,2})\2(\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?$/.exec(s);
  if (nd) {
    const a = Number(nd[1]);
    const b = Number(nd[3]);
    let y = Number(nd[4]);
    if (y < 100) y += y < 30 ? 2000 : 1900;
    let dayFirst = DAY_FIRST || nd[2] === '.';
    if (a > 12) dayFirst = true;
    if (b > 12) dayFirst = false;
    const [m, d] = dayFirst ? [b, a] : [a, b];
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      const v = partsToSerial(y, m, d, Number(nd[5] ?? 0), Number(nd[6] ?? 0));
      const fmt = dayFirst ? (nd[2] === '.' ? 'dd.mm.yyyy' : 'd/m/yyyy') : 'm/d/yyyy';
      return { value: v, format: nd[5] ? `${fmt} h:mm` : fmt };
    }
  }
  // month names: 15-Mar-2024, 15 March 2024, Mar 15, 2024, March 2024
  const md1 = /^(\d{1,2})[\s-]([A-Za-z]{3,9})[\s-,]*(\d{2,4})?$/.exec(s);
  const md2 = /^([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s*(\d{4})?$/.exec(s);
  if (md1 || md2) {
    const day = Number(md1 ? md1[1] : md2![2]);
    const mon = monthFromName(md1 ? md1[2] : md2![1]);
    let y = Number((md1 ? md1[3] : md2![3]) ?? new Date().getFullYear());
    if (y < 100) y += 2000;
    if (mon >= 1 && day >= 1 && day <= 31) {
      return { value: partsToSerial(y, mon, day), format: md1 ? (md1[3] ? 'd-mmm-yyyy' : 'd-mmm') : 'mmm d, yyyy' };
    }
  }
  return { value: raw };
}
