import { SHORT_DATE } from '../format/numfmt';

export type NumCategory = 'general' | 'number' | 'currency' | 'accounting' | 'date' | 'time' | 'percent' | 'fraction' | 'scientific' | 'text' | 'special' | 'custom';

export const CATEGORIES: { id: NumCategory; label: string; hint: string }[] = [
  { id: 'general', label: 'General', hint: 'No specific number format.' },
  { id: 'number', label: 'Number', hint: 'For general display of numbers.' },
  { id: 'currency', label: 'Currency', hint: 'For general monetary values.' },
  { id: 'accounting', label: 'Accounting', hint: 'Lines up currency symbols and decimal points in a column.' },
  { id: 'date', label: 'Date', hint: 'Displays date serial numbers as dates.' },
  { id: 'time', label: 'Time', hint: 'Displays date serial numbers as times.' },
  { id: 'percent', label: 'Percentage', hint: 'Multiplies the cell value by 100 and shows a % symbol.' },
  { id: 'fraction', label: 'Fraction', hint: 'Shows numbers as fractions.' },
  { id: 'scientific', label: 'Scientific', hint: 'Shows numbers in exponential notation.' },
  { id: 'text', label: 'Text', hint: 'The cell is treated as text, even when a number is typed.' },
  { id: 'special', label: 'Special', hint: 'Postal codes, phone numbers and more.' },
  { id: 'custom', label: 'Custom', hint: 'Type your own format code.' },
];

export const CURRENCIES = [
  { symbol: '$', label: '$ Dollar' },
  { symbol: '€', label: '€ Euro' },
  { symbol: '£', label: '£ Pound' },
  { symbol: '¥', label: '¥ Yen / Yuan' },
  { symbol: '₹', label: '₹ Rupee' },
  { symbol: 'CHF', label: 'CHF Swiss franc' },
  { symbol: 'kr', label: 'kr Krona / Krone' },
  { symbol: 'R$', label: 'R$ Real' },
  { symbol: '₩', label: '₩ Won' },
  { symbol: 'zł', label: 'zł Złoty' },
  { symbol: '₽', label: '₽ Ruble' },
  { symbol: '₺', label: '₺ Lira' },
  { symbol: 'R', label: 'R Rand' },
  { symbol: 'A$', label: 'A$ Australian dollar' },
  { symbol: 'C$', label: 'C$ Canadian dollar' },
];

export const NEGATIVE_STYLES = ['minus', 'red', 'parens', 'redParens'] as const;
export type NegativeStyle = (typeof NEGATIVE_STYLES)[number];

export const DATE_FORMATS = [SHORT_DATE, 'dddd, mmmm d, yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'mmmm d, yyyy', 'yyyy-mm-dd', 'd/m/yyyy', 'dd.mm.yyyy', 'mmm d', 'mmmm yyyy', `${SHORT_DATE} h:mm`, 'yyyy-mm-dd h:mm:ss', 'ddd, mmm d'];
export const TIME_FORMATS = ['h:mm', 'h:mm AM/PM', 'h:mm:ss', 'h:mm:ss AM/PM', '[h]:mm:ss', 'mm:ss', 'mm:ss.0', '[h]:mm'];
export const FRACTION_FORMATS = [
  ['# ?/?', 'Up to one digit (1/4)'],
  ['# ??/??', 'Up to two digits (21/25)'],
  ['# ???/???', 'Up to three digits (312/943)'],
  ['# ?/2', 'As halves (1/2)'],
  ['# ?/4', 'As quarters (2/4)'],
  ['# ?/8', 'As eighths (4/8)'],
  ['# ??/16', 'As sixteenths (8/16)'],
  ['# ?/10', 'As tenths (3/10)'],
  ['# ??/100', 'As hundredths (30/100)'],
] as const;
export const SPECIAL_FORMATS = [
  ['00000', 'Zip code'],
  ['00000-0000', 'Zip code + 4'],
  ['[<=9999999]###-####;(###) ###-####', 'Phone number'],
  ['000-00-0000', 'Social security number'],
  ['+0 000 000 0000', 'International phone'],
] as const;

export const CUSTOM_PRESETS = [
  'General',
  '0',
  '0.00',
  '#,##0',
  '#,##0.00',
  '#,##0;-#,##0',
  '#,##0;[Red]-#,##0',
  '#,##0.00;[Red](#,##0.00)',
  '0%',
  '0.00%',
  '0.00E+00',
  '# ?/?',
  'm/d/yyyy',
  'd-mmm-yy',
  'mmm-yy',
  'h:mm AM/PM',
  '[h]:mm:ss',
  '@',
  '0.0,,"M"',
  '0.0,"K"',
  '[Green]▲ 0.0%;[Red]▼ 0.0%;0.0%',
  '"Qty: "0',
  '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0',
  ';;;',
];

function quoteSym(sym: string): string {
  return sym === '$' ? '$' : `"${sym}"`;
}

export interface NumOptions {
  decimals: number;
  thousands: boolean;
  negative: NegativeStyle;
  symbol: string;
  symbolAfter?: boolean;
}

export function buildFormat(cat: NumCategory, o: NumOptions, pick?: string): string {
  const dec = o.decimals > 0 ? '.' + '0'.repeat(o.decimals) : '';
  switch (cat) {
    case 'general':
      return 'General';
    case 'number':
    case 'currency': {
      const core = `${o.thousands || cat === 'currency' ? '#,##0' : '0'}${dec}`;
      const pos = cat === 'currency' ? (o.symbolAfter ? `${core} ${quoteSym(o.symbol)}` : `${quoteSym(o.symbol)}${core}`) : core;
      switch (o.negative) {
        case 'minus':
          return pos;
        case 'red':
          return `${pos};[Red]-${pos}`;
        case 'parens':
          return `${pos}_);(${pos})`;
        case 'redParens':
          return `${pos}_);[Red](${pos})`;
      }
      return pos;
    }
    case 'accounting': {
      const s = quoteSym(o.symbol);
      return `_(${s}* #,##0${dec}_);_(${s}* (#,##0${dec});_(${s}* "-"${o.decimals ? '??' : ''}_);_(@_)`;
    }
    case 'percent':
      return `0${dec}%`;
    case 'scientific':
      return `0${dec}E+00`;
    case 'text':
      return '@';
    default:
      return pick ?? 'General';
  }
}

/** Best-effort reverse mapping of a format code to a dialog category. */
export function categorize(fmt: string | undefined): { cat: NumCategory; opts: NumOptions } {
  const f = fmt ?? 'General';
  const opts: NumOptions = { decimals: 2, thousands: false, negative: 'minus', symbol: '$' };
  const decs = (/\.(0+)/.exec(f)?.[1] ?? '').length;
  opts.decimals = decs;
  opts.thousands = f.includes('#,##0');
  if (/\[Red\]\(/i.test(f)) opts.negative = 'redParens';
  else if (/\[Red\]/i.test(f)) opts.negative = 'red';
  else if (/;\(/.test(f)) opts.negative = 'parens';
  const sym = /"([^"]+)"|(\$)/.exec(f);
  if (sym) opts.symbol = sym[1] ?? sym[2];
  if (/^general$/i.test(f)) return { cat: 'general', opts };
  if (f === '@') return { cat: 'text', opts };
  if (/^_\(/.test(f)) return { cat: 'accounting', opts };
  if (DATE_FORMATS.includes(f)) return { cat: 'date', opts };
  if (TIME_FORMATS.includes(f)) return { cat: 'time', opts };
  if (FRACTION_FORMATS.some(([x]) => x === f)) return { cat: 'fraction', opts };
  if (SPECIAL_FORMATS.some(([x]) => x === f)) return { cat: 'special', opts };
  if (/^0(\.0+)?%$/.test(f)) return { cat: 'percent', opts };
  if (/^0(\.0+)?E\+00$/.test(f)) return { cat: 'scientific', opts };
  if (/^(\$|"[^"]+")#,##0/.test(f) || /#,##0(\.0+)? "[^"]+"/.test(f)) return { cat: 'currency', opts };
  if (/^(#,##0|0)(\.0+)?(_\))?(;.*)?$/.test(f)) return { cat: 'number', opts };
  return { cat: 'custom', opts };
}

/** Increase/decrease decimals of an existing format (ribbon buttons). */
export function stepDecimals(fmt: string | undefined, dir: 1 | -1, sample?: number): string {
  const f = fmt && !/^general$/i.test(fmt) ? fmt : null;
  if (!f) {
    // General: derive from the sample's current decimals
    const s = sample !== undefined && Number.isFinite(sample) ? String(parseFloat(sample.toPrecision(10))) : '0';
    const d = (s.split('.')[1] ?? '').length;
    const n = Math.max(0, d + dir);
    return n ? `0.${'0'.repeat(n)}` : '0';
  }
  return f.split(';').map((sec) => {
    if (/[dmyhs]/i.test(sec.replace(/"[^"]*"/g, '')) && !/0/.test(sec)) return sec;
    if (dir > 0) {
      if (/0\.0*/.test(sec)) return sec.replace(/(0\.0*)/, '$10');
      return sec.replace(/(0)(?!.*0)/, '$1.0');
    }
    if (/0\.0(?!0)/.test(sec)) return sec.replace(/0\.0(?!0)/, '0');
    return sec.replace(/(0\.0*)0/, '$1');
  }).join(';');
}
