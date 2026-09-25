/**
 * Number formats in OpenDocument are structured "data styles" (number:number-style, number:date-style…)
 * instead of format codes. These helpers translate Excel-style format codes to data styles and back.
 */
import { sectionsFor, type FormatSection, type FormatTok } from '../format/numfmt';

export const NS_NUMBER = 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0';
const NS_STYLE = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';
const NS_FO = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const CURRENCY = new Set(['$', '€', '£', '¥', '₹', '¢', '₩', '₽', '₺', '₫', '₪', '฿', '₴', '₦']);

type Kind = 'number' | 'percentage' | 'currency' | 'date' | 'time' | 'text';

interface Built {
  kind: Kind;
  body: string;
  /** Extra attributes for the style element. */
  attrs?: string;
}

/* ============================================================ writing */

function text(v: string): string {
  return v ? `<number:text>${esc(v)}</number:text>` : '';
}

function numberSection(sec: FormatSection): Built {
  const toks = sec.toks;
  const hasDigits = toks.some((t) => t.t === 'digit' || t.t === 'general');
  const expAt = toks.findIndex((t) => t.t === 'exp');
  const slashAt = toks.findIndex((t) => t.t === 'slash');
  let kind: Kind = 'number';
  const out: string[] = [];
  let pending = '';
  const flush = () => {
    out.push(text(pending));
    pending = '';
  };
  let numberDone = false;
  // number part: from the first to the last digit-ish token
  const isNum = (t: FormatTok) => t.t === 'digit' || t.t === 'point' || t.t === 'comma' || t.t === 'exp' || t.t === 'slash' || t.t === 'general' || (t.t === 'lit' && t.v.startsWith('\u0000'));
  const first = toks.findIndex(isNum);
  let last = -1;
  toks.forEach((t, i) => {
    if (isNum(t)) last = i;
  });
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (i >= first && i <= last && first >= 0) {
      if (!numberDone) {
        flush();
        out.push(numberElement(toks.slice(first, last + 1), expAt >= 0, slashAt >= 0));
        numberDone = true;
      }
      continue;
    }
    if (t.t === 'percent') {
      kind = 'percentage';
      pending += '%';
    } else if (t.t === 'lit' && CURRENCY.has(t.v) && hasDigits && kind !== 'percentage') {
      flush();
      kind = 'currency';
      out.push(`<number:currency-symbol>${esc(t.v)}</number:currency-symbol>`);
    } else if (t.t === 'lit') pending += t.v;
  }
  flush();
  return { kind, body: out.join('') };
}

function numberElement(toks: FormatTok[], sci: boolean, frac: boolean): string {
  if (toks.some((t) => t.t === 'general')) return '<number:number number:min-integer-digits="1"/>';
  if (frac) {
    // "# ?/?", "# ??/16", "?/?"
    const slash = toks.findIndex((t) => t.t === 'slash');
    const before = toks.slice(0, slash);
    const after = toks.slice(slash + 1);
    const gap = before.findIndex((t, i) => t.t === 'lit' && i > 0);
    const intToks = gap >= 0 ? before.slice(0, gap) : [];
    const numToks = gap >= 0 ? before.slice(gap + 1) : before;
    const minInt = intToks.filter((t) => t.t === 'digit' && t.v === '0').length;
    const numDigits = numToks.filter((t) => t.t === 'digit').length;
    const fixed = after.find((t) => t.t === 'lit' && t.v.startsWith('\u0000'));
    const denDigits = after.filter((t) => t.t === 'digit').length;
    const den = fixed && fixed.t === 'lit' ? ` number:denominator-value="${fixed.v.slice(1)}"` : ` number:min-denominator-digits="${Math.max(1, denDigits)}"`;
    return `<number:fraction${gap >= 0 ? ` number:min-integer-digits="${minInt}"` : ''} number:min-numerator-digits="${Math.max(1, numDigits)}"${den}/>`;
  }
  const expAt = toks.findIndex((t) => t.t === 'exp');
  const mant = sci ? toks.slice(0, expAt) : toks;
  const pointAt = mant.findIndex((t) => t.t === 'point');
  const intToks = pointAt >= 0 ? mant.slice(0, pointAt) : mant;
  const decToks = pointAt >= 0 ? mant.slice(pointAt + 1) : [];
  const intDigits = intToks.filter((t) => t.t === 'digit');
  const minInt = intDigits.filter((t) => t.t === 'digit' && t.v === '0').length;
  // a comma between integer digits groups thousands; commas after them divide by 1000 each
  const lastDigit = intToks.map((t) => t.t).lastIndexOf('digit');
  const grouping = intToks.some((t, i) => t.t === 'comma' && i < lastDigit);
  const trailing = intToks.slice(lastDigit + 1).filter((t) => t.t === 'comma').length + (pointAt < 0 ? 0 : decToks.slice(decToks.map((t) => t.t).lastIndexOf('digit') + 1).filter((t) => t.t === 'comma').length);
  const decDigits = decToks.filter((t) => t.t === 'digit');
  const minDec = decDigits.filter((t) => t.t === 'digit' && t.v === '0').length;
  const at: string[] = [];
  at.push(`number:decimal-places="${decDigits.length}"`);
  at.push(`number:min-decimal-places="${minDec}"`);
  at.push(`number:min-integer-digits="${minInt}"`);
  if (grouping) at.push('number:grouping="true"');
  if (sci) {
    const expDigits = toks.slice(expAt + 1).filter((t) => t.t === 'digit').length;
    const exp = toks[expAt];
    at.push(`number:min-exponent-digits="${Math.max(1, expDigits)}"`);
    if (exp.t === 'exp' && exp.sign === '+') at.push('number:forced-exponent-sign="true"');
    return `<number:scientific-number ${at.join(' ')}/>`;
  }
  if (trailing) at.push(`number:display-factor="${1000 ** trailing}"`);
  // literals between integer digits (phone numbers, IDs…) become embedded text
  let embedded = '';
  let digitsRight = intDigits.length;
  let lit = '';
  for (const t of intToks) {
    if (t.t === 'digit') {
      if (lit) {
        embedded += `<number:embedded-text number:position="${digitsRight}">${esc(lit)}</number:embedded-text>`;
        lit = '';
      }
      digitsRight--;
    } else if (t.t === 'lit' && digitsRight > 0 && digitsRight < intDigits.length) lit += t.v;
  }
  return embedded ? `<number:number ${at.join(' ')}>${embedded}</number:number>` : `<number:number ${at.join(' ')}/>`;
}

function dateSection(sec: FormatSection): Built {
  const out: string[] = [];
  let pending = '';
  let dateLike = false;
  let elapsed = false;
  const flush = () => {
    out.push(text(pending));
    pending = '';
  };
  const el = (name: string, long: boolean, extra = '') => {
    flush();
    out.push(`<number:${name}${long ? ' number:style="long"' : ''}${extra}/>`);
  };
  const toks = sec.toks;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === 'date') {
      const v = t.v;
      if (v[0] === 'y') {
        dateLike = true;
        el('year', v.length > 2);
      } else if (v[0] === 'm') {
        dateLike = true;
        if (v.length >= 3) el('month', v.length === 4, ' number:textual="true"');
        else el('month', v.length === 2);
      } else if (v[0] === 'd') {
        dateLike = true;
        if (v.length >= 3) el('day-of-week', v.length >= 4);
        else el('day', v.length === 2);
      } else if (v[0] === 'h') el('hours', v.length >= 2);
      else if (v[0] === 'M') el('minutes', v.length >= 2);
      else if (v[0] === 's') {
        const sub = toks[i + 1];
        el('seconds', v.length >= 2, sub?.t === 'subsec' ? ` number:decimal-places="${sub.len}"` : '');
      }
    } else if (t.t === 'elapsed') {
      elapsed = true;
      el(t.v === 'h' ? 'hours' : t.v === 'm' ? 'minutes' : 'seconds', t.len >= 2);
    } else if (t.t === 'ampm') el('am-pm', false);
    else if (t.t === 'lit') pending += t.v;
    else if (t.t === 'point') pending += '.';
    else if (t.t === 'comma') pending += ',';
  }
  flush();
  return { kind: dateLike ? 'date' : 'time', body: out.join(''), attrs: elapsed ? ' number:truncate-on-overflow="false"' : '' };
}

function textSection(sec: FormatSection): Built {
  let body = '';
  let pending = '';
  for (const t of sec.toks) {
    if (t.t === 'text') {
      body += text(pending) + '<number:text-content/>';
      pending = '';
    } else if (t.t === 'lit') pending += t.v;
    else if (t.t === 'comma') pending += ',';
  }
  return { kind: 'text', body: body + text(pending) };
}

function buildSection(sec: FormatSection): Built {
  if (sec.isText) return textSection(sec);
  if (sec.isDate) return dateSection(sec);
  return numberSection(sec);
}

const TAG: Record<Kind, string> = {
  number: 'number:number-style',
  percentage: 'number:percentage-style',
  currency: 'number:currency-style',
  date: 'number:date-style',
  time: 'number:time-style',
  text: 'number:text-style',
};

function styleXml(name: string, b: Built, color: string | undefined, maps: string, volatile: boolean): string {
  const tag = TAG[b.kind];
  const props = color ? `<style:text-properties fo:color="${color}"/>` : '';
  return `<${tag} style:name="${name}"${volatile ? ' style:volatile="true"' : ''}${b.attrs ?? ''}>${props}${b.body}${maps}</${tag}>`;
}

function condText(c: { op: string; v: number }): string {
  const op = c.op === '=' ? '=' : c.op === '<>' ? '!=' : c.op;
  return `value()${op}${c.v}`;
}

/**
 * Data style XML (the named main style plus any conditional sub-styles) for an Excel format code.
 * Returns an empty string for "General", which needs no data style.
 */
export function numFmtToOds(code: string, name: string): string {
  if (!code || /^general$/i.test(code.trim())) return '';
  const sections = sectionsFor(code);
  if (sections.length === 1) {
    const s = sections[0];
    return styleXml(name, buildSection(s), s.color, '', false);
  }
  // Excel: positive;negative;zero;text — ODF: the last applicable style carries the conditions
  const explicit = sections.some((s) => s.cond);
  const built = sections.map(buildSection);
  let conds: string[];
  let mainIdx: number;
  if (explicit) {
    mainIdx = sections.length - 1;
    conds = sections.slice(0, mainIdx).map((s, i) => (s.cond ? condText(s.cond) : i === 0 ? 'value()>=0' : 'value()<0'));
  } else if (sections.length === 2) {
    mainIdx = 1;
    conds = ['value()>=0'];
  } else if (sections.length === 3) {
    mainIdx = 2;
    conds = ['value()>0', 'value()<0'];
  } else {
    mainIdx = 3;
    conds = ['value()>0', 'value()<0', 'value()=0'];
  }
  let xml = '';
  let maps = '';
  conds.forEach((c, i) => {
    const sub = `${name}P${i}`;
    xml += styleXml(sub, built[i], sections[i].color, '', true);
    maps += `<style:map style:condition="${esc(c)}" style:apply-style-name="${sub}"/>`;
  });
  return xml + styleXml(name, built[mainIdx], sections[mainIdx].color, maps, false);
}

/** The OpenDocument value type of a number shown with a format code (from its first section). */
export function odsValueType(code: string | undefined): 'float' | 'percentage' | 'currency' | 'date' | 'time' {
  if (!code || /^general$/i.test(code.trim())) return 'float';
  let type = valueTypes.get(code);
  if (!type) {
    const sec = sectionsFor(code)[0];
    if (!sec || sec.isText) type = 'float';
    else if (sec.isDate) type = sec.toks.some((t) => t.t === 'date' && /^[ymd]/.test(t.v)) ? 'date' : 'time';
    else {
      const kind = numberSection(sec).kind;
      type = kind === 'percentage' || kind === 'currency' ? kind : 'float';
    }
    valueTypes.set(code, type);
  }
  return type;
}

const valueTypes = new Map<string, ReturnType<typeof odsValueType>>();

/* ============================================================ reading */

function kids(el: Element): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
  return out;
}

const num = (el: Element, name: string) => {
  const v = el.getAttributeNS(NS_NUMBER, name);
  return v === null || v === '' ? undefined : Number(v);
};

const NAMED_COLORS: Record<string, string> = { '#000000': 'Black', '#ffffff': 'White', '#ff0000': 'Red', '#00ff00': 'Green', '#0000ff': 'Blue', '#ffff00': 'Yellow', '#ff00ff': 'Magenta', '#00ffff': 'Cyan' };

/** An Excel literal: quoted unless it's a character Excel shows as-is. */
function lit(v: string): string {
  if (!v) return '';
  if (/^[\s$\-+/():!^&'~{}<>=,.€£¥₹%]*$/.test(v) && !v.includes('"')) return v.replace(/%/g, '\\%');
  if (!v.includes('"')) return `"${v}"`;
  return [...v].map((c) => `\\${c}`).join('');
}

function intPattern(minInt: number, grouping: boolean): string {
  let p = minInt > 0 ? '0'.repeat(minInt) : '#';
  if (grouping) {
    p = p.padStart(4, '#');
    p = `${p.slice(0, p.length - 3)},${p.slice(-3)}`;
  }
  return p;
}

function numberCode(el: Element): string {
  if (el.localName === 'scientific-number') {
    const minInt = num(el, 'min-integer-digits') ?? 1;
    const dec = num(el, 'decimal-places') ?? 2;
    const minDec = num(el, 'min-decimal-places') ?? dec;
    const exp = num(el, 'min-exponent-digits') ?? 2;
    const sign = el.getAttributeNS(NS_NUMBER, 'forced-exponent-sign') === 'false' ? '-' : '+';
    return `${'0'.repeat(Math.max(1, minInt))}${dec ? `.${'0'.repeat(minDec)}${'#'.repeat(Math.max(0, dec - minDec))}` : ''}E${sign}${'0'.repeat(exp)}`;
  }
  if (el.localName === 'fraction') {
    const minInt = num(el, 'min-integer-digits');
    const numDigits = num(el, 'min-numerator-digits') ?? 1;
    const denValue = num(el, 'denominator-value');
    const denDigits = num(el, 'min-denominator-digits') ?? 1;
    const intPart = minInt === undefined ? '' : `${minInt > 0 ? '0'.repeat(minInt) : '#'} `;
    return `${intPart}${'?'.repeat(numDigits)}/${denValue ? String(denValue) : '?'.repeat(denDigits)}`;
  }
  // number:number (LibreOffice may leave out decimal-places when min-decimal-places says the same)
  const currency = (el.parentNode as Element | null)?.localName === 'currency-style';
  const dec = num(el, 'decimal-places') ?? num(el, 'min-decimal-places') ?? (currency ? 2 : undefined);
  const minInt = num(el, 'min-integer-digits') ?? 1;
  const grouping = el.getAttributeNS(NS_NUMBER, 'grouping') === 'true';
  if (dec === undefined && !grouping && minInt <= 1 && !kids(el).length) return 'General';
  const minDec = num(el, 'min-decimal-places') ?? dec ?? 0;
  let intP = intPattern(minInt, grouping);
  // embedded text sits a number of digits left of the decimal point
  const embedded = kids(el).filter((k) => k.localName === 'embedded-text');
  if (embedded.length) {
    const digits = [...intP.replace(/,/g, '')];
    while (digits.length < Math.max(...embedded.map((e) => num(e, 'position') ?? 0))) digits.unshift('#');
    const parts: string[] = [];
    digits.forEach((d, i) => {
      const pos = digits.length - i;
      for (const e of embedded) if ((num(e, 'position') ?? -1) === pos && i > 0) parts.push(lit(e.textContent ?? ''));
      parts.push(d);
    });
    intP = parts.join('');
  }
  const factor = num(el, 'display-factor');
  const commas = factor && factor > 1 ? ','.repeat(Math.round(Math.log10(factor) / 3)) : '';
  const decP = dec ? `.${'0'.repeat(Math.min(minDec, dec))}${'#'.repeat(Math.max(0, dec - minDec))}` : '';
  return `${intP}${decP}${commas}`;
}

function dateCode(el: Element): string {
  const long = (k: Element) => k.getAttributeNS(NS_NUMBER, 'style') === 'long';
  const elapsed = el.getAttributeNS(NS_NUMBER, 'truncate-on-overflow') === 'false';
  let firstTime = true;
  let out = '';
  for (const k of kids(el)) {
    switch (k.localName) {
      case 'year':
        out += long(k) ? 'yyyy' : 'yy';
        break;
      case 'month':
        out += k.getAttributeNS(NS_NUMBER, 'textual') === 'true' ? (long(k) ? 'mmmm' : 'mmm') : long(k) ? 'mm' : 'm';
        break;
      case 'day':
        out += long(k) ? 'dd' : 'd';
        break;
      case 'day-of-week':
        out += long(k) ? 'dddd' : 'ddd';
        break;
      case 'hours':
      case 'minutes':
      case 'seconds': {
        const ch = k.localName[0];
        let code = long(k) ? ch + ch : ch;
        if (elapsed && firstTime) code = `[${code}]`;
        firstTime = false;
        out += code;
        if (k.localName === 'seconds') {
          const d = num(k, 'decimal-places');
          if (d) out += `.${'0'.repeat(d)}`;
        }
        break;
      }
      case 'am-pm':
        out += 'AM/PM';
        break;
      case 'text':
        out += lit(k.textContent ?? '');
        break;
      default:
        break;
    }
  }
  return out;
}

function currencyCode(sym: string): string {
  if (!sym) return '$';
  if (CURRENCY.has(sym)) return sym;
  return `[$${sym}]`;
}

/** The code of one data style without its conditional maps. */
function sectionCode(el: Element): string {
  const color = kids(el)
    .find((k) => k.localName === 'text-properties')
    ?.getAttributeNS(NS_FO, 'color')
    ?.toLowerCase();
  const prefix = color && NAMED_COLORS[color] ? `[${NAMED_COLORS[color]}]` : '';
  switch (el.localName) {
    case 'date-style':
    case 'time-style':
      return prefix + dateCode(el);
    case 'boolean-style':
      return 'General';
    default:
      break;
  }
  let out = '';
  for (const k of kids(el)) {
    switch (k.localName) {
      case 'number':
      case 'scientific-number':
      case 'fraction':
        out += numberCode(k);
        break;
      case 'text': {
        const t = k.textContent ?? '';
        // in percentage styles the % sign is a text element but must stay a real percent
        out += el.localName === 'percentage-style' && t.includes('%') ? t.split('%').map(lit).join('%') : lit(t);
        break;
      }
      case 'currency-symbol':
        out += currencyCode((k.textContent ?? '').trim());
        break;
      case 'text-content':
        out += '@';
        break;
      default:
        break;
    }
  }
  if (el.localName === 'percentage-style' && !out.includes('%')) out += '%';
  return prefix + out;
}

/**
 * Excel format code for an ODF data style (undefined for the default "General" format).
 * `lookup` resolves the sub-styles referenced by style:map conditions.
 */
export function odsToNumFmt(el: Element, lookup: (name: string) => Element | undefined): string | undefined {
  const main = sectionCode(el);
  const maps = kids(el)
    .filter((k) => k.localName === 'map')
    .map((m) => ({ cond: (m.getAttributeNS(NS_STYLE, 'condition') ?? '').replace(/\s/g, ''), style: lookup(m.getAttributeNS(NS_STYLE, 'apply-style-name') ?? '') }))
    .filter((m): m is { cond: string; style: Element } => !!m.style);
  let code: string;
  if (!maps.length) code = main;
  else {
    const conds = maps.map((m) => m.cond);
    const subs = maps.map((m) => sectionCode(m.style));
    const same = (a: string[]) => a.length === conds.length && a.every((c, i) => c === conds[i]);
    if (same(['value()>=0'])) code = `${subs[0]};${main}`;
    else if (same(['value()>0', 'value()<0'])) code = `${subs[0]};${subs[1]};${main}`;
    else if (same(['value()>0', 'value()<0', 'value()=0'])) code = `${subs[0]};${subs[1]};${subs[2]};${main}`;
    else {
      const bracket = (c: string) => {
        const m = /^value\(\)(<=|>=|!=|<|>|=)(-?[\d.]+)$/.exec(c);
        return m ? `[${m[1] === '!=' ? '<>' : m[1]}${m[2]}]` : '';
      };
      code = [...subs.map((s, i) => bracket(conds[i]) + s), main].join(';');
    }
  }
  return !code || /^general$/i.test(code) ? undefined : code;
}
