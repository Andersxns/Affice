/**
 * Formulas in OpenDocument files use OpenFormula syntax: `of:=SUM([.A1:.B2];[$Sheet2.C3])`.
 * These helpers translate between it and the Excel syntax the Sheets engine uses.
 */
import { tokenize, type Token } from '../engine/parser';
import { quoteSheet } from '../model/address';

/** Excel functions LibreOffice stores with a COM.MICROSOFT. prefix (not part of ODF 1.2). */
const MS_PREFIXED = new Set(
  `AGGREGATE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST
CONCAT CONFIDENCE.NORM CONFIDENCE.T COVARIANCE.P COVARIANCE.S ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FLOOR.MATH FLOOR.PRECISE
FORECAST.LINEAR FORECAST.ETS.ADD FORECAST.ETS.MULT FORECAST.ETS.PI.ADD FORECAST.ETS.PI.MULT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT.ADD FORECAST.ETS.STAT.MULT
GAMMA.DIST GAMMA.INV GAMMALN.PRECISE HYPGEOM.DIST IFS ISO.CEILING LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS MODE.MULT MODE.SNGL NEGBINOM.DIST NETWORKDAYS.INTL
NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC POISSON.DIST QUARTILE.EXC QUARTILE.INC RANK.AVG RANK.EQ
STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TEXTJOIN VAR.P VAR.S WEIBULL.DIST WORKDAY.INTL Z.TEST XLOOKUP XMATCH LET LAMBDA
SEQUENCE RANDARRAY UNIQUE SORT SORTBY FILTER TEXTBEFORE TEXTAFTER TEXTSPLIT VSTACK HSTACK TOCOL TOROW WRAPROWS WRAPCOLS TAKE DROP CHOOSEROWS CHOOSECOLS EXPAND
MAP REDUCE SCAN MAKEARRAY BYROW BYCOL ISOMITTED VALUETOTEXT ARRAYTOTEXT REGEXTEST REGEXEXTRACT REGEXREPLACE GROUPBY PIVOTBY FILTERXML ENCODEURL`.split(/\s+/),
);

/** Excel functions whose OpenFormula name differs. */
const XL_TO_OF: Record<string, string> = {
  FORMULATEXT: 'FORMULA',
  'SKEW.P': 'SKEWP',
  CEILING: 'COM.MICROSOFT.CEILING',
  FLOOR: 'COM.MICROSOFT.FLOOR',
  CHIDIST: 'LEGACY.CHIDIST',
  CHIINV: 'LEGACY.CHIINV',
  CHITEST: 'LEGACY.CHITEST',
  FDIST: 'LEGACY.FDIST',
  FINV: 'LEGACY.FINV',
  NORMSDIST: 'LEGACY.NORMSDIST',
  NORMSINV: 'LEGACY.NORMSINV',
  TDIST: 'LEGACY.TDIST',
};
const OF_TO_XL: Record<string, string> = { FORMULA: 'FORMULATEXT', SKEWP: 'SKEW.P' };
const OF_PREFIXES = ['COM.MICROSOFT.', 'LEGACY.', 'ORG.OPENOFFICE.', 'ORG.LIBREOFFICE.', '_XLFN._XLWS.', '_XLFN.', '_XLWS.'];

/** A sheet name as written inside an OpenFormula reference (`$Sheet1` or `$'My sheet'`). */
export function odsSheet(name: string): string {
  return `$${/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`}`;
}

function fnToOds(name: string): string {
  const up = name.toUpperCase().replace(/^_XLFN\.(_XLWS\.)?/, '');
  return XL_TO_OF[up] ?? (MS_PREFIXED.has(up) ? `COM.MICROSOFT.${up}` : name);
}

function fnFromOds(name: string): string {
  let up = name.toUpperCase();
  for (const p of OF_PREFIXES)
    if (up.startsWith(p)) {
      up = up.slice(p.length);
      break;
    }
  return OF_TO_XL[up] ?? up;
}

/** Excel-syntax formula (without "=") → OpenFormula (with the "of:=" namespace prefix). */
export function toOdsFormula(f: string): string {
  const toks = tokenize(f);
  let out = '';
  let braces = 0;
  // for each open parenthesis: does it start a function's argument list?
  const parens: boolean[] = [];
  const sheetOf = (t: Token) => (t.sheet !== undefined ? odsSheet(t.sheet) : '');
  const significant = (from: number, step: 1 | -1) => {
    for (let k = from; k >= 0 && k < toks.length; k += step) if (toks[k].type !== 'ws') return toks[k];
    return undefined;
  };
  const RANGE_END = new Set(['ref', 'colrange', 'rowrange', 'name', 'rparen']);
  const RANGE_START = new Set(['ref', 'colrange', 'rowrange', 'name', 'lparen']);
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    switch (t.type) {
      case 'lparen':
        parens.push(significant(i - 1, -1)?.type === 'func');
        out += '(';
        break;
      case 'rparen':
        parens.pop();
        out += ')';
        break;
      case 'ws': {
        // a space between two references is Excel's intersection operator
        const prev = significant(i - 1, -1);
        const next = significant(i + 1, 1);
        out += prev && next && RANGE_END.has(prev.type) && RANGE_START.has(next.type) && braces === 0 ? '!' : t.text;
        break;
      }
      case 'ref': {
        const end = toks[i + 1]?.type === 'colon' && toks[i + 2]?.type === 'ref' ? toks[i + 2] : undefined;
        if (end) i += 2;
        out += `[${sheetOf(t)}.${t.refText}${end ? `:${end.sheet !== undefined ? sheetOf(end) : ''}.${end.refText}` : ''}]`;
        break;
      }
      case 'colrange':
      case 'rowrange': {
        const [a, b] = (t.refText ?? '').split(':');
        out += `[${sheetOf(t)}.${a}:.${b}]`;
        break;
      }
      case 'func':
        out += fnToOds(t.text);
        break;
      case 'bool':
        out += `${t.text.toUpperCase()}()`;
        break;
      case 'name':
        // sheet-scoped names keep only their name
        out += t.refText ?? t.text;
        break;
      case 'sep':
        // inside plain parentheses a comma is the union operator
        out += braces > 0 || parens[parens.length - 1] ? ';' : '~';
        break;
      case 'semi':
        out += braces > 0 ? '|' : ';';
        break;
      case 'lbrace':
        braces++;
        out += '{';
        break;
      case 'rbrace':
        braces = Math.max(0, braces - 1);
        out += '}';
        break;
      case 'op':
        if (t.text !== '@') out += t.text; // implicit intersection has no OpenFormula form
        break;
      case 'hash':
        break; // spill references (A1#) point at the formula cell itself
      default:
        out += t.text;
    }
  }
  return `of:=${out}`;
}

/** Converts the inside of an OpenFormula reference ("$Sheet1.A1:.B2") to Excel syntax ("Sheet1!A1:B2"). */
export function odsRefToExcel(inner: string): string {
  if (/#REF!/i.test(inner)) return '#REF!';
  const parts: string[] = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i];
    if (ch === "'") {
      if (q && inner[i + 1] === "'") {
        cur += "''";
        i++;
        continue;
      }
      q = !q;
    }
    if (ch === ':' && !q) {
      parts.push(cur);
      cur = '';
    } else cur += ch;
  }
  parts.push(cur);
  let sheet: string | undefined;
  const cells = parts.map((p, n) => {
    // the cell part follows the last dot outside quotes
    let dot = -1;
    let inQ = false;
    for (let i = 0; i < p.length; i++) {
      if (p[i] === "'") inQ = !inQ;
      else if (p[i] === '.' && !inQ) dot = i;
    }
    const sheetPart = dot >= 0 ? p.slice(0, dot).replace(/^\$/, '') : '';
    if (sheetPart && n === 0) sheet = sheetPart.startsWith("'") ? sheetPart.slice(1, -1).replace(/''/g, "'") : sheetPart;
    return dot >= 0 ? p.slice(dot + 1) : p;
  });
  const range = cells.join(':');
  return sheet !== undefined ? `${quoteSheet(sheet)}!${range}` : range;
}

/** OpenFormula (with or without a namespace prefix such as "of:=") → Excel syntax without "=". */
export function fromOdsFormula(src: string): string {
  let s = src.trim();
  const ns = /^([a-z][\w-]*):=?/i.exec(s);
  if (ns) {
    s = s.slice(ns[0].length);
    // formulas written by Excel into ODS files can use Excel syntax already
    if (ns[1].toLowerCase() === 'msoxl') return s.replace(/^=/, '');
  }
  s = s.replace(/^=/, '');
  let out = '';
  let braces = 0;
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < s.length && !(s[j] === '"' && s[j + 1] !== '"')) j += s[j] === '"' ? 2 : 1;
      out += s.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === '[') {
      let j = i + 1;
      let q = false;
      while (j < s.length && (q || s[j] !== ']')) {
        if (s[j] === "'") q = !q;
        j++;
      }
      out += odsRefToExcel(s.slice(i + 1, j));
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const err = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|GETTING_DATA)/i.exec(s.slice(i));
      if (err) {
        out += err[0].toUpperCase();
        i += err[0].length;
        continue;
      }
    }
    if (ch === '!') {
      // intersection
      out += ' ';
      i++;
      continue;
    }
    if (ch === '{') braces++;
    else if (ch === '}') braces = Math.max(0, braces - 1);
    // ';' separates arguments, and columns inside array constants: both are ',' in Excel
    if (ch === ';') {
      out += ',';
      i++;
      continue;
    }
    if (ch === '|' && braces > 0) {
      out += ';';
      i++;
      continue;
    }
    if (ch === '~') {
      out += ',';
      i++;
      continue;
    }
    const id = /^[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*/.exec(s.slice(i));
    if (id) {
      let j = i + id[0].length;
      while (s[j] === ' ') j++;
      if (s[j] === '(') {
        const up = id[0].toUpperCase();
        // TRUE() and FALSE() are functions in OpenFormula, constants in Excel
        if ((up === 'TRUE' || up === 'FALSE') && /^\(\s*\)/.test(s.slice(j))) {
          out += up;
          i = j + s.slice(j).indexOf(')') + 1;
          continue;
        }
        out += fnFromOds(id[0]);
        i += id[0].length;
        continue;
      }
      out += id[0];
      i += id[0].length;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

/** Splits a list of ODF range addresses at spaces outside quoted sheet names. */
export function splitRangeList(list: string): string[] {
  const out: string[] = [];
  let cur = '';
  let q = false;
  for (const ch of list) {
    if (ch === "'") q = !q;
    if (/\s/.test(ch) && !q) {
      if (cur) out.push(cur);
      cur = '';
    } else cur += ch;
  }
  if (cur) out.push(cur);
  return out;
}

/** A cell range as an ODF range address ("Sheet1.A1:Sheet1.B5"), used by filters, conditional formats… */
export function odsRangeAddress(sheet: string, r1: number, c1: number, r2: number, c2: number, abs = false): string {
  const sh = abs ? odsSheet(sheet) : odsSheet(sheet).slice(1);
  const cell = (r: number, c: number) => `${abs ? '$' : ''}${colLetters(c)}${abs ? '$' : ''}${r + 1}`;
  const a = `${sh}.${cell(r1, c1)}`;
  return r1 === r2 && c1 === c2 ? a : `${a}:${sh}.${cell(r2, c2)}`;
}

function colLetters(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
