import { ERR, isErr, asErr, isLambda, isMatrix, isRef, type Scalar, type Value } from '../values';
import { deref, type EvalContext } from '../evaluator';
import { addrName } from '../../model/address';
import { fn, num, str } from './helpers';

const I = 'Information' as const;

fn('ISBLANK', I, 1, 1, 'e', 'value', 'Checks whether a reference is to an empty cell.', ([v]) => v === null);
fn('ISERR', I, 1, 1, 'e', 'value', 'Checks whether a value is an error other than #N/A.', ([v]) => isErr(v) && asErr(v).error !== '#N/A');
fn('ISERROR', I, 1, 1, 'e', 'value', 'Checks whether a value is an error.', ([v]) => isErr(v));
fn('ISNA', I, 1, 1, 'e', 'value', 'Checks whether a value is #N/A.', ([v]) => isErr(v) && asErr(v).error === '#N/A');
fn('ISLOGICAL', I, 1, 1, 'e', 'value', 'Checks whether a value is a logical value (TRUE or FALSE).', ([v]) => typeof v === 'boolean');
fn('ISNUMBER', I, 1, 1, 'e', 'value', 'Checks whether a value is a number.', ([v]) => typeof v === 'number');
fn('ISTEXT', I, 1, 1, 'e', 'value', 'Checks whether a value is text.', ([v]) => typeof v === 'string');
fn('ISNONTEXT', I, 1, 1, 'e', 'value', 'Checks whether a value is not text (blank cells are not text).', ([v]) => typeof v !== 'string');
fn('ISEVEN', I, 1, 1, 'v', 'number', 'Returns TRUE if the number is even.', ([v]) => Math.trunc(num(v)) % 2 === 0);
fn('ISODD', I, 1, 1, 'v', 'number', 'Returns TRUE if the number is odd.', ([v]) => Math.abs(Math.trunc(num(v)) % 2) === 1);
fn('ISREF', I, 1, 1, 'r', 'value', 'Checks whether a value is a reference.', ([v]: Value[]) => isRef(v));
fn('ISFORMULA', I, 1, 1, 'r', 'reference', 'Checks whether a reference is to a cell containing a formula.', ([v]: Value[], ctx) => {
  if (!isRef(v)) throw ERR.VALUE;
  return ctx.formulaText(v.sheet, v.r1, v.c1) !== undefined;
});

fn('N', I, 1, 1, 'v', 'value', 'Converts a non-number value to a number, dates to serial numbers, TRUE to 1, anything else to 0.', ([v]) => {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  return 0;
});
fn('NA', I, 0, 0, 'v', '', 'Returns the error value #N/A.', () => ERR.NA);

const ERROR_CODES: Record<string, number> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#GETTING_DATA': 8,
  '#SPILL!': 9,
  '#CALC!': 14,
  '#CIRC!': 4,
};

fn('ERROR.TYPE', I, 1, 1, 'e', 'error_val', 'Returns a number matching an error value.', ([v]) => {
  if (!isErr(v)) throw ERR.NA;
  return ERROR_CODES[asErr(v).error] ?? ERR.NA;
});

fn('TYPE', I, 1, 1, 'r', 'value', 'Returns an integer representing the data type of a value: 1 number, 2 text, 4 logical, 16 error, 64 array, 128 lambda.', ([v]: Value[], ctx) => {
  if (isLambda(v)) return 128;
  if (isMatrix(v)) return 64;
  if (isRef(v) && (v.r1 !== v.r2 || v.c1 !== v.c2)) return 64;
  const x = deref(v, ctx) as Scalar;
  if (typeof x === 'number' || x === null) return 1;
  if (typeof x === 'string') return 2;
  if (typeof x === 'boolean') return 4;
  return 16;
});

fn('SHEET', I, 0, 1, 'r', '[value]', 'Returns the sheet number of the referenced sheet.', (args: Value[], ctx: EvalContext) => {
  const ids = ctx.sheetIds();
  if (!args.length || args[0] === null) return ids.indexOf(ctx.sheet) + 1;
  const v = args[0];
  if (isRef(v)) return ids.indexOf(v.sheet) + 1;
  if (typeof v === 'string') {
    const id = ctx.sheetId(v);
    if (id === undefined) throw ERR.NA;
    return ids.indexOf(id) + 1;
  }
  throw ERR.VALUE;
});

fn('SHEETS', I, 0, 1, 'r', '[reference]', 'Returns the number of sheets in a reference (or in the workbook).', (args: Value[], ctx: EvalContext) => {
  if (args.length && isRef(args[0])) return 1;
  return ctx.sheetIds().length;
});

function formatCode(fmt: string | undefined): string {
  const f = (fmt ?? 'General').toLowerCase();
  if (f === 'general') return 'G';
  const neg = f.includes(';') && (f.split(';')[1].includes('(') ? '()' : f.split(';')[1].includes('[red]') ? '-' : '');
  const dec = (f.split(';')[0].split('.')[1] ?? '').replace(/[^0#]/g, '').length;
  if (/[dmy]/.test(f) && !/[hs]/.test(f)) {
    if (/d.*m.*y|m.*d.*y/.test(f)) return 'D1';
    if (/d.*m/.test(f)) return 'D2';
    return 'D3';
  }
  if (/[hs]/.test(f)) return /am\/pm/.test(f) ? 'D7' : 'D9';
  if (f.includes('%')) return `P${dec}`;
  if (f.includes('e+') || f.includes('e-')) return `S${dec}`;
  if (/[$€£¥]/.test(f)) return `C${dec}${neg === '()' ? '()' : neg === '-' ? '-' : ''}`;
  if (f.includes(',')) return `,${dec}`;
  return `F${dec}`;
}

fn('CELL', I, 1, 2, 'vr', 'info_type, [reference]', 'Returns information about the formatting, location, or contents of a cell.', (args: Value[], ctx: EvalContext) => {
  const info = str(args[0] as Scalar).toLowerCase();
  let ref = args[1];
  if (ref === undefined || ref === null) ref = { kind: 'ref', sheet: ctx.sheet, r1: ctx.row, c1: ctx.col, r2: ctx.row, c2: ctx.col };
  if (!isRef(ref)) throw ERR.VALUE;
  const v = ctx.cell(ref.sheet, ref.r1, ref.c1);
  switch (info) {
    case 'address':
      return (ref.sheet !== ctx.sheet ? `'${ctx.sheetName(ref.sheet)}'!` : '') + addrName(ref.r1, ref.c1, true, true);
    case 'row':
      return ref.r1 + 1;
    case 'col':
      return ref.c1 + 1;
    case 'contents':
      return v === null ? 0 : v;
    case 'type':
      return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v';
    case 'filename':
      return ctx.sheetName(ref.sheet) ?? '';
    case 'format':
      return formatCode(ctx.cellFormat?.(ref.sheet, ref.r1, ref.c1));
    case 'color':
      return /\[red\]/i.test(ctx.cellFormat?.(ref.sheet, ref.r1, ref.c1) ?? '') ? 1 : 0;
    case 'parentheses':
      return /\(/.test((ctx.cellFormat?.(ref.sheet, ref.r1, ref.c1) ?? '').split(';')[0]) ? 1 : 0;
    case 'prefix':
      return typeof v === 'string' ? "'" : '';
    case 'protect':
      return 1;
    case 'width':
      return 10;
  }
  throw ERR.VALUE;
});

fn('INFO', I, 1, 1, 'v', 'type_text', 'Returns information about the current operating environment.', ([t]) => {
  switch (str(t).toLowerCase()) {
    case 'directory':
      return '';
    case 'numfile':
      return 1;
    case 'origin':
      return '$A:$A$1';
    case 'osversion':
      return typeof navigator !== 'undefined' ? navigator.platform : 'Affice';
    case 'recalc':
      return 'Automatic';
    case 'release':
      return '16.0';
    case 'system':
      return typeof navigator !== 'undefined' && /mac/i.test(navigator.platform) ? 'mac' : 'pcdos';
  }
  throw ERR.VALUE;
});
