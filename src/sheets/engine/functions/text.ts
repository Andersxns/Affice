import { ERR, isErr, asErr, matrix, type Scalar, type Value } from '../values';
import { formatValue, parseInput } from '../../format/numfmt';
import { bool, check, flat, fn, grid, int, num, optBool, optInt, optStr, roundTo, str } from './helpers';

const T = 'Text' as const;

/* -------------------------------------------------------- windows-1252 */

const CP1252: Record<number, number> = {
  128: 0x20ac,
  130: 0x201a,
  131: 0x0192,
  132: 0x201e,
  133: 0x2026,
  134: 0x2020,
  135: 0x2021,
  136: 0x02c6,
  137: 0x2030,
  138: 0x0160,
  139: 0x2039,
  140: 0x0152,
  142: 0x017d,
  145: 0x2018,
  146: 0x2019,
  147: 0x201c,
  148: 0x201d,
  149: 0x2022,
  150: 0x2013,
  151: 0x2014,
  152: 0x02dc,
  153: 0x2122,
  154: 0x0161,
  155: 0x203a,
  156: 0x0153,
  158: 0x017e,
  159: 0x0178,
};
const CP1252_REV = new Map(Object.entries(CP1252).map(([k, v]) => [v, Number(k)]));

/* --------------------------------------------------------------- basics */

fn('LEN', T, 1, 1, 'v', 'text', 'Returns the number of characters in a text string.', ([t]) => [...str(t)].length);
fn('LENB', T, 1, 1, 'v', 'text', 'Returns the number of bytes used to represent the characters in a text string.', ([t]) => [...str(t)].length);

fn('LEFT', T, 1, 2, 'v', 'text, [num_chars]', 'Returns the leftmost characters of a text value.', (args) => {
  const n = optInt(args, 1, 1);
  check(n >= 0, ERR.VALUE);
  return [...str(args[0])].slice(0, n).join('');
});
fn('RIGHT', T, 1, 2, 'v', 'text, [num_chars]', 'Returns the rightmost characters of a text value.', (args) => {
  const n = optInt(args, 1, 1);
  check(n >= 0, ERR.VALUE);
  const chars = [...str(args[0])];
  return n === 0 ? '' : chars.slice(Math.max(0, chars.length - n)).join('');
});
fn('MID', T, 3, 3, 'v', 'text, start_num, num_chars', 'Returns a specific number of characters from a text string starting at the position you specify.', ([t, s, n]) => {
  const start = int(s);
  const len = int(n);
  check(start >= 1 && len >= 0, ERR.VALUE);
  return [...str(t)].slice(start - 1, start - 1 + len).join('');
});
for (const [b, base] of [
  ['LEFTB', 'LEFT'],
  ['RIGHTB', 'RIGHT'],
  ['MIDB', 'MID'],
] as const) {
  fn(b, T, base === 'MID' ? 3 : 1, base === 'MID' ? 3 : 2, 'v', base === 'MID' ? 'text, start_num, num_bytes' : 'text, [num_bytes]', `Like ${base}, counting bytes.`, (args) => {
    const chars = [...str(args[0])];
    if (base === 'LEFT') return chars.slice(0, optInt(args, 1, 1)).join('');
    if (base === 'RIGHT') {
      const n = optInt(args, 1, 1);
      return n === 0 ? '' : chars.slice(Math.max(0, chars.length - n)).join('');
    }
    const start = int(args[1]);
    check(start >= 1, ERR.VALUE);
    return chars.slice(start - 1, start - 1 + int(args[2])).join('');
  });
}

fn('UPPER', T, 1, 1, 'v', 'text', 'Converts text to uppercase.', ([t]) => str(t).toUpperCase());
fn('LOWER', T, 1, 1, 'v', 'text', 'Converts text to lowercase.', ([t]) => str(t).toLowerCase());
fn('PROPER', T, 1, 1, 'v', 'text', 'Capitalizes the first letter in each word of a text value.', ([t]) => {
  let prevLetter = false;
  let out = '';
  for (const ch of str(t)) {
    const isLetter = ch.toLowerCase() !== ch.toUpperCase();
    out += isLetter ? (prevLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
    prevLetter = isLetter;
  }
  return out;
});
fn('TRIM', T, 1, 1, 'v', 'text', 'Removes spaces from text, leaving single spaces between words.', ([t]) => str(t).replace(/ +/g, ' ').replace(/^ | $/g, ''));
fn('CLEAN', T, 1, 1, 'v', 'text', 'Removes all nonprintable characters from text.', ([t]) => str(t).replace(/[\x00-\x1f]/g, ''));
fn('EXACT', T, 2, 2, 'v', 'text1, text2', 'Checks whether two text values are exactly the same (case-sensitive).', ([a, b]) => str(a) === str(b));
fn('REPT', T, 2, 2, 'v', 'text, number_times', 'Repeats text a given number of times.', ([t, n]) => {
  const k = int(n);
  const s = str(t);
  check(k >= 0 && s.length * k <= 32767, ERR.VALUE);
  return s.repeat(k);
});
fn('T', T, 1, 1, 'e', 'value', 'Returns the text if the value is text, otherwise an empty string.', ([v]) => (isErr(v) ? asErr(v) : typeof v === 'string' ? v : ''));

fn('CHAR', T, 1, 1, 'v', 'number', 'Returns the character specified by the code number (1-255).', ([n]) => {
  const k = int(n);
  check(k >= 1 && k <= 255, ERR.VALUE);
  return String.fromCharCode(CP1252[k] ?? k);
});
fn('CODE', T, 1, 1, 'v', 'text', 'Returns a numeric code for the first character in a text string.', ([t]) => {
  const s = str(t);
  check(s.length > 0, ERR.VALUE);
  const cp = s.charCodeAt(0);
  if (cp < 256) return cp;
  return CP1252_REV.get(cp) ?? 63;
});
fn('UNICHAR', T, 1, 1, 'v', 'number', 'Returns the Unicode character for a number.', ([n]) => {
  const k = int(n);
  check(k >= 1 && k <= 0x10ffff, ERR.VALUE);
  return String.fromCodePoint(k);
});
fn('UNICODE', T, 1, 1, 'v', 'text', 'Returns the number (code point) of the first character of the text.', ([t]) => {
  const s = str(t);
  check(s.length > 0, ERR.VALUE);
  return s.codePointAt(0)!;
});

/* ------------------------------------------------------------ joining */

fn('CONCATENATE', T, 1, 255, 'v', 'text1, [text2], …', 'Joins several text items into one text item.', (args) => {
  let out = '';
  for (const a of args) out += str(a);
  check(out.length <= 32767, ERR.VALUE);
  return out;
});

fn('CONCAT', T, 1, 254, 'r', 'text1, [text2], …', 'Joins a list or range of text strings.', (args: Value[], ctx) => {
  let out = '';
  for (const a of args) {
    for (const v of flat(a, ctx)) {
      if (isErr(v)) throw asErr(v);
      out += str(v);
    }
  }
  check(out.length <= 32767, ERR.VALUE);
  return out;
});

fn('TEXTJOIN', T, 3, 252, 'r', 'delimiter, ignore_empty, text1, [text2], …', 'Joins text from ranges with a delimiter, optionally skipping empty values.', (args: Value[], ctx) => {
  const delims = flat(args[0], ctx).map((d) => {
    if (isErr(d)) throw asErr(d);
    return str(d);
  });
  const ignore = bool(flat(args[1], ctx)[0]);
  const parts: string[] = [];
  for (const a of args.slice(2)) {
    for (const v of flat(a, ctx)) {
      if (isErr(v)) throw asErr(v);
      const s = str(v);
      if (ignore && s === '') continue;
      parts.push(s);
    }
  }
  let out = parts[0] ?? '';
  for (let i = 1; i < parts.length; i++) out += (delims.length ? delims[(i - 1) % delims.length] : '') + parts[i];
  check(out.length <= 32767, ERR.VALUE);
  return out;
});

/* ------------------------------------------------------------- search */

fn('FIND', T, 2, 3, 'v', 'find_text, within_text, [start_num]', 'Finds one text value within another (case-sensitive).', (args) => {
  const find = str(args[0]);
  const within = str(args[1]);
  const start = optInt(args, 2, 1);
  check(start >= 1 && start <= within.length + (find === '' ? 1 : 0), ERR.VALUE);
  const i = within.indexOf(find, start - 1);
  if (i < 0) throw ERR.VALUE;
  return i + 1;
});
fn('FINDB', T, 2, 3, 'v', 'find_text, within_text, [start_num]', 'Finds one text value within another (case-sensitive).', (args) => {
  const within = str(args[1]);
  const i = within.indexOf(str(args[0]), optInt(args, 2, 1) - 1);
  if (i < 0) throw ERR.VALUE;
  return i + 1;
});

function wildcardSearchRegex(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) re += '\\' + pattern[++i];
    else if (ch === '*') re += '[\\s\\S]*?';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(re, 'i');
}

fn('SEARCH', T, 2, 3, 'v', 'find_text, within_text, [start_num]', 'Finds one text value within another (not case-sensitive, supports wildcards).', (args) => {
  const find = str(args[0]);
  const within = str(args[1]);
  const start = optInt(args, 2, 1);
  check(start >= 1 && start <= within.length + (find === '' ? 1 : 0), ERR.VALUE);
  const m = wildcardSearchRegex(find).exec(within.slice(start - 1));
  if (!m) throw ERR.VALUE;
  return m.index + start;
});
fn('SEARCHB', T, 2, 3, 'v', 'find_text, within_text, [start_num]', 'Finds one text value within another (not case-sensitive).', (args) => {
  const within = str(args[1]);
  const start = optInt(args, 2, 1);
  const m = wildcardSearchRegex(str(args[0])).exec(within.slice(start - 1));
  if (!m) throw ERR.VALUE;
  return m.index + start;
});

fn('SUBSTITUTE', T, 3, 4, 'v', 'text, old_text, new_text, [instance_num]', 'Substitutes new text for old text in a text string.', (args) => {
  const text = str(args[0]);
  const old = str(args[1]);
  const rep = str(args[2]);
  if (old === '') return text;
  if (args.length > 3) {
    const n = int(args[3]);
    check(n >= 1, ERR.VALUE);
    let idx = -1;
    for (let k = 0; k < n; k++) {
      idx = text.indexOf(old, idx + 1);
      if (idx < 0) return text;
    }
    return text.slice(0, idx) + rep + text.slice(idx + old.length);
  }
  return text.split(old).join(rep);
});

fn('REPLACE', T, 4, 4, 'v', 'old_text, start_num, num_chars, new_text', 'Replaces characters within text.', ([t, s, n, r]) => {
  const chars = [...str(t)];
  const start = int(s);
  const len = int(n);
  check(start >= 1 && len >= 0, ERR.VALUE);
  return chars.slice(0, start - 1).join('') + str(r) + chars.slice(start - 1 + len).join('');
});
fn('REPLACEB', T, 4, 4, 'v', 'old_text, start_num, num_bytes, new_text', 'Replaces characters within text.', ([t, s, n, r]) => {
  const chars = [...str(t)];
  const start = int(s);
  return chars.slice(0, start - 1).join('') + str(r) + chars.slice(start - 1 + int(n)).join('');
});

/* ---------------------------------------------------------- conversion */

fn('TEXT', T, 2, 2, 'v', 'value, format_text', 'Formats a number and converts it to text using a number format code.', ([v, f]) => {
  const fmt = str(f);
  let val: Scalar = v;
  if (typeof v === 'string') {
    const p = parseInput(v);
    if (typeof p.value === 'number') val = p.value;
  }
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
  return formatValue(val ?? 0, fmt).text;
});

fn('VALUE', T, 1, 1, 'v', 'text', 'Converts a text argument that looks like a number (or date/time) to a number.', ([t]) => {
  if (typeof t === 'number') return t;
  if (typeof t === 'boolean') throw ERR.VALUE;
  const s = str(t).trim();
  if (s === '') return 0;
  const p = parseInput(s);
  if (typeof p.value !== 'number') throw ERR.VALUE;
  return p.value;
});

fn('NUMBERVALUE', T, 1, 3, 'v', 'text, [decimal_separator], [group_separator]', 'Converts text to a number in a locale-independent way.', (args) => {
  let s = str(args[0]).replace(/\s+/g, '');
  const dec = optStr(args, 1, '.').charAt(0) || '.';
  const grp = optStr(args, 2, ',').charAt(0) || ',';
  if (s === '') return 0;
  let pct = 0;
  while (s.endsWith('%')) {
    pct++;
    s = s.slice(0, -1);
  }
  s = s.split(grp).join('');
  const parts = s.split(dec);
  if (parts.length > 2) throw ERR.VALUE;
  const normal = parts.join('.');
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(normal)) throw ERR.VALUE;
  return parseFloat(normal) / Math.pow(100, pct);
});

function groupThousands(intPart: string): string {
  return intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

fn('FIXED', T, 1, 3, 'v', 'number, [decimals], [no_commas]', 'Formats a number as text with a fixed number of decimals.', (args) => {
  const n = num(args[0]);
  const d = optInt(args, 1, 2);
  const noCommas = optBool(args, 2, false);
  check(d <= 127, ERR.VALUE);
  const r = roundTo(n, d);
  let s = Math.abs(r).toFixed(Math.max(0, d));
  if (!noCommas) {
    const [i, f] = s.split('.');
    s = groupThousands(i) + (f !== undefined ? '.' + f : '');
  }
  return (r < 0 ? '-' : '') + s;
});

fn('DOLLAR', T, 1, 2, 'v', 'number, [decimals]', 'Converts a number to text using currency format.', (args) => {
  const n = num(args[0]);
  const d = optInt(args, 1, 2);
  const r = roundTo(n, d);
  const [i, f] = Math.abs(r).toFixed(Math.max(0, d)).split('.');
  const body = '$' + groupThousands(i) + (f !== undefined ? '.' + f : '');
  return r < 0 ? `(${body})` : body;
});

fn('VALUETOTEXT', T, 1, 2, 'e', 'value, [format]', 'Returns text from any specified value.', (args) => {
  const v = args[0] as Scalar;
  const strict = optInt(args, 1, 0) === 1;
  return valueText(v, strict);
});

function valueText(v: Scalar, strict: boolean): string {
  if (isErr(v)) return asErr(v).error;
  if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v;
  return str(v);
}

fn('ARRAYTOTEXT', T, 1, 2, 'rv', 'array, [format]', 'Returns an array of text values from any specified range.', (args: Value[], ctx) => {
  const strict = args.length > 1 && num(args[1] as Scalar) === 1;
  const g = grid(args[0], ctx);
  if (!strict) return g.map((row) => row.map((v) => valueText(v, false)).join(', ')).join(', ');
  return '{' + g.map((row) => row.map((v) => valueText(v, true)).join(',')).join(';') + '}';
});

for (const name of ['ASC', 'DBCS', 'JIS', 'PHONETIC']) {
  fn(name, T, 1, 1, 'v', 'text', 'Returns the text unchanged (full/half-width conversion for East Asian text).', ([t]) => str(t));
}

/* ------------------------------------------------ TEXTBEFORE / AFTER */

function delimiterList(v: Value, ctx: Parameters<typeof flat>[1]): string[] {
  return flat(v, ctx).map((d) => {
    if (isErr(d)) throw asErr(d);
    return str(d);
  });
}

function findDelim(text: string, delims: string[], from: number, ci: boolean): [number, number] | null {
  const hay = ci ? text.toLowerCase() : text;
  let best: [number, number] | null = null;
  for (const d of delims) {
    const needle = ci ? d.toLowerCase() : d;
    const i = hay.indexOf(needle, from);
    if (i >= 0 && (!best || i < best[0])) best = [i, d.length];
  }
  return best;
}

function findDelimRev(text: string, delims: string[], before: number, ci: boolean): [number, number] | null {
  const hay = ci ? text.toLowerCase() : text;
  let best: [number, number] | null = null;
  for (const d of delims) {
    const needle = ci ? d.toLowerCase() : d;
    const i = hay.lastIndexOf(needle, before);
    if (i >= 0 && (!best || i > best[0])) best = [i, d.length];
  }
  return best;
}

function textAround(args: Value[], ctx: Parameters<typeof flat>[1], after: boolean): Scalar {
  const text = str(args[0] as Scalar);
  const delims = delimiterList(args[1], ctx);
  const inst = args.length > 2 && args[2] !== null ? int(args[2] as Scalar) : 1;
  const ci = args.length > 3 && args[3] !== null ? int(args[3] as Scalar) === 1 : false;
  const matchEnd = args.length > 4 && args[4] !== null ? bool(args[4] as Scalar) : false;
  const notFound = args.length > 5 ? (args[5] as Scalar) : ERR.NA;
  if (inst === 0 || Math.abs(inst) > text.length + 1) throw ERR.VALUE;
  if (delims.some((d) => d === '')) return after ? (inst > 0 ? text : '') : inst > 0 ? '' : text;
  let pos: [number, number] | null = null;
  if (inst > 0) {
    let from = 0;
    for (let k = 0; k < inst; k++) {
      pos = findDelim(text, delims, from, ci);
      if (!pos) {
        if (matchEnd && k === inst - 1) pos = [text.length, 0];
        break;
      }
      from = pos[0] + Math.max(1, pos[1]);
    }
  } else {
    let before = text.length;
    for (let k = 0; k < -inst; k++) {
      pos = findDelimRev(text, delims, before, ci);
      if (!pos) {
        if (matchEnd && k === -inst - 1) pos = [0, 0];
        break;
      }
      before = pos[0] - 1;
      if (before < 0 && k < -inst - 1) {
        pos = null;
        break;
      }
    }
  }
  if (!pos) return notFound;
  return after ? text.slice(pos[0] + pos[1]) : text.slice(0, pos[0]);
}

fn('TEXTBEFORE', T, 2, 6, 'vrvvvv', 'text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]', 'Returns text that occurs before a given character or string.', (args, ctx) => textAround(args, ctx, false));
fn('TEXTAFTER', T, 2, 6, 'vrvvvv', 'text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]', 'Returns text that occurs after a given character or string.', (args, ctx) => textAround(args, ctx, true));

fn('TEXTSPLIT', T, 2, 6, 'vrrvvv', 'text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]', 'Splits text into rows and columns using delimiters.', (args: Value[], ctx) => {
  const text = str(args[0] as Scalar);
  const colD = args[1] === null ? [] : delimiterList(args[1], ctx).filter((d) => d !== '');
  const rowD = args.length > 2 && args[2] !== null ? delimiterList(args[2], ctx).filter((d) => d !== '') : [];
  const ignore = args.length > 3 && args[3] !== null ? bool(args[3] as Scalar) : false;
  const ci = args.length > 4 && args[4] !== null ? int(args[4] as Scalar) === 1 : false;
  const pad: Scalar = args.length > 5 ? (args[5] as Scalar) : ERR.NA;
  if (!colD.length && !rowD.length) throw ERR.VALUE;
  const split = (s: string, ds: string[]): string[] => {
    if (!ds.length) return [s];
    const esc = ds.map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|');
    const parts = s.split(new RegExp(esc, ci ? 'i' : ''));
    return ignore ? parts.filter((p) => p !== '') : parts;
  };
  const rows = split(text, rowD).map((r) => split(r, colD));
  const width = Math.max(1, ...rows.map((r) => r.length));
  return matrix(rows.map((r) => [...r, ...Array(width - r.length).fill(pad)] as Scalar[]));
});

/* -------------------------------------------------------------- regex */

function makeRegex(pattern: string, flags: string): RegExp {
  try {
    return new RegExp(pattern, flags);
  } catch {
    throw ERR.VALUE;
  }
}

fn('REGEXTEST', T, 2, 3, 'v', 'text, pattern, [case_sensitivity]', 'Checks whether any part of the text matches a regular expression.', (args) => {
  const ci = optInt(args, 2, 0) === 1;
  return makeRegex(str(args[1]), ci ? 'i' : '').test(str(args[0]));
});

fn('REGEXEXTRACT', T, 2, 4, 'v', 'text, pattern, [return_mode], [case_sensitivity]', 'Extracts text matching a regular expression (0 first match, 1 all matches, 2 capture groups).', (args) => {
  const text = str(args[0]);
  const mode = optInt(args, 2, 0);
  const ci = optInt(args, 3, 0) === 1;
  if (mode === 1) {
    const re = makeRegex(str(args[1]), ci ? 'gi' : 'g');
    const all = [...text.matchAll(re)].map((m) => [m[0]] as Scalar[]);
    if (!all.length) throw ERR.NA;
    return matrix(all);
  }
  const m = makeRegex(str(args[1]), ci ? 'i' : '').exec(text);
  if (!m) throw ERR.NA;
  if (mode === 2) {
    const groups = m.slice(1);
    if (!groups.length) return m[0];
    return matrix([groups.map((g) => g ?? '')]);
  }
  return m[0];
});

fn('REGEXREPLACE', T, 3, 5, 'v', 'text, pattern, replacement, [occurrence], [case_sensitivity]', 'Replaces text matching a regular expression.', (args) => {
  const text = str(args[0]);
  const rep = str(args[2]);
  const occ = optInt(args, 3, 0);
  const ci = optInt(args, 4, 0) === 1;
  const re = makeRegex(str(args[1]), ci ? 'gi' : 'g');
  if (occ === 0) return text.replace(re, rep);
  const ms = [...text.matchAll(re)];
  const target = occ > 0 ? ms[occ - 1] : ms[ms.length + occ];
  if (!target || target.index === undefined) return text;
  const single = makeRegex(str(args[1]), ci ? 'i' : '');
  return text.slice(0, target.index) + target[0].replace(single, rep) + text.slice(target.index + target[0].length);
});

/* ---------------------------------------------------------------- web */

fn('ENCODEURL', 'Web', 1, 1, 'v', 'text', 'Returns a URL-encoded string.', ([t]) => encodeURIComponent(str(t)));
