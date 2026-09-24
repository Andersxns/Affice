import { colIndex, MAX_COLS, MAX_ROWS } from '../model/address';

/* ================================================================ lexer */

export type TokType =
  | 'num'
  | 'str'
  | 'bool'
  | 'err'
  | 'ref'
  | 'colrange'
  | 'rowrange'
  | 'func'
  | 'name'
  | 'op'
  | 'lparen'
  | 'rparen'
  | 'sep'
  | 'semi'
  | 'lbrace'
  | 'rbrace'
  | 'colon'
  | 'hash'
  | 'ws'
  | 'bad';

export interface Token {
  type: TokType;
  text: string;
  start: number;
  end: number;
  /** For refs: sheet prefix as written (unquoted) */
  sheet?: string;
  /** For refs: text of the reference part without the sheet prefix */
  refText?: string;
}

const ERRORS = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#SPILL!', '#CALC!', '#GETTING_DATA', '#CIRC!'];

const CELL_RE = /^\$?[A-Za-z]{1,3}\$?\d{1,7}/;
const COLRANGE_RE = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w(])/;
const ROWRANGE_RE = /^\$?\d{1,7}:\$?\d{1,7}(?![\w(.])/;
const IDENT_RE = /^[A-Za-z_\\À-￿][\w.À-￿]*/;
const SHEET_UNQUOTED_RE = /^([A-Za-z_À-￿][\w.À-￿]*)!/;

function validCell(s: string): boolean {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(s);
  if (!m) return false;
  return colIndex(m[1]) < MAX_COLS && Number(m[2]) >= 1 && Number(m[2]) <= MAX_ROWS;
}

export function tokenize(src: string): Token[] {
  const toks: Token[] = [];
  let i = 0;
  const n = src.length;
  const push = (type: TokType, start: number, end: number, extra: Partial<Token> = {}) => toks.push({ type, text: src.slice(start, end), start, end, ...extra });
  const prevSignificant = () => {
    for (let k = toks.length - 1; k >= 0; k--) if (toks[k].type !== 'ws') return toks[k];
    return undefined;
  };
  while (i < n) {
    const ch = src[i];
    const rest = src.slice(i);
    if (/\s/.test(ch)) {
      let j = i;
      while (j < n && /\s/.test(src[j])) j++;
      push('ws', i, j);
      i = j;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') j += 2;
          else break;
        } else j++;
      }
      push(j < n ? 'str' : 'bad', i, Math.min(n, j + 1));
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const e = ERRORS.find((x) => rest.toUpperCase().startsWith(x));
      if (e) {
        push('err', i, i + e.length);
        i += e.length;
        continue;
      }
      const prev = prevSignificant();
      if (prev && (prev.type === 'ref' || prev.type === 'rparen')) {
        push('hash', i, i + 1);
        i++;
        continue;
      }
      push('bad', i, i + 1);
      i++;
      continue;
    }
    // numbers
    const num = /^(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/.exec(rest);
    if (num && !ROWRANGE_RE.test(rest)) {
      push('num', i, i + num[0].length);
      i += num[0].length;
      continue;
    }
    // sheet-qualified refs: 'Sheet name'!A1 or Sheet1!A1:B2
    let sheet: string | undefined;
    let j = i;
    if (ch === "'") {
      let k = i + 1;
      let name = '';
      while (k < n) {
        if (src[k] === "'") {
          if (src[k + 1] === "'") {
            name += "'";
            k += 2;
            continue;
          }
          break;
        }
        name += src[k++];
      }
      if (src[k] === "'" && src[k + 1] === '!') {
        sheet = name;
        j = k + 2;
      }
    } else {
      const sm = SHEET_UNQUOTED_RE.exec(rest);
      if (sm) {
        sheet = sm[1];
        j = i + sm[0].length;
      }
    }
    const after = src.slice(j);
    const cr = COLRANGE_RE.exec(after);
    if (cr) {
      push('colrange', i, j + cr[0].length, { sheet, refText: cr[0] });
      i = j + cr[0].length;
      continue;
    }
    const rr = ROWRANGE_RE.exec(after);
    if (rr) {
      push('rowrange', i, j + rr[0].length, { sheet, refText: rr[0] });
      i = j + rr[0].length;
      continue;
    }
    const cm = CELL_RE.exec(after);
    if (cm && !/^[\w.(]/.test(after.slice(cm[0].length)) && validCell(cm[0])) {
      push('ref', i, j + cm[0].length, { sheet, refText: cm[0] });
      i = j + cm[0].length;
      continue;
    }
    if (sheet !== undefined) {
      // Sheet-scoped name, or a #REF! target
      const id = IDENT_RE.exec(after);
      if (id) {
        push('name', i, j + id[0].length, { sheet, refText: id[0] });
        i = j + id[0].length;
        continue;
      }
      if (after.toUpperCase().startsWith('#REF!')) {
        push('err', i, j + 5);
        i = j + 5;
        continue;
      }
    }
    const id = IDENT_RE.exec(rest);
    if (id) {
      const word = id[0];
      let k = i + word.length;
      while (k < n && src[k] === ' ') k++;
      if (src[k] === '(') {
        push('func', i, i + word.length);
        i += word.length;
        continue;
      }
      const up = word.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') push('bool', i, i + word.length);
      else push('name', i, i + word.length);
      i += word.length;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      push('op', i, i + 2);
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%@'.includes(ch)) {
      push('op', i, i + 1);
      i++;
      continue;
    }
    const simple: Record<string, TokType> = { '(': 'lparen', ')': 'rparen', ',': 'sep', ';': 'semi', '{': 'lbrace', '}': 'rbrace', ':': 'colon' };
    if (simple[ch]) {
      push(simple[ch], i, i + 1);
      i++;
      continue;
    }
    push('bad', i, i + 1);
    i++;
  }
  return toks;
}

/* ================================================================== AST */

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: string }
  | { t: 'missing' }
  | { t: 'ref'; sheet?: string; r: number; c: number; ra: boolean; ca: boolean }
  | { t: 'range'; sheet?: string; r1: number; c1: number; r2: number; c2: number; ra1: boolean; ca1: boolean; ra2: boolean; ca2: boolean; kind: 'cells' | 'cols' | 'rows' }
  | { t: 'spill'; ref: Node }
  | { t: 'name'; name: string; sheet?: string }
  | { t: 'func'; name: string; args: Node[] }
  | { t: 'call'; callee: Node; args: Node[] }
  | { t: 'unary'; op: string; arg: Node }
  | { t: 'postfix'; op: '%'; arg: Node }
  | { t: 'bin'; op: string; left: Node; right: Node }
  | { t: 'array'; rows: Node[][] };

export class FormulaError extends Error {
  constructor(
    message: string,
    public pos = 0,
  ) {
    super(message);
  }
}

function parseCellRef(text: string): { r: number; c: number; ra: boolean; ca: boolean } {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(text)!;
  return { ca: m[1] === '$', c: colIndex(m[2]), ra: m[3] === '$', r: Number(m[4]) - 1 };
}

const BIN_PREC: Record<string, number> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
};
const PREC_UNARY = 6;
const PREC_POSTFIX = 7;

export function parseFormula(src: string): Node {
  const toks = tokenize(src).filter((t) => t.type !== 'ws');
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (type: TokType, what: string) => {
    const t = next();
    if (!t || t.type !== type) throw new FormulaError(`Expected ${what}`, t?.start ?? src.length);
    return t;
  };

  const refNode = (t: Token): Node => {
    const txt = t.refText!;
    if (t.type === 'ref') return { t: 'ref', sheet: t.sheet, ...parseCellRef(txt) };
    if (t.type === 'colrange') {
      const [a, b] = txt.split(':');
      const c1 = colIndex(a.replace('$', ''));
      const c2 = colIndex(b.replace('$', ''));
      return { t: 'range', sheet: t.sheet, r1: 0, c1: Math.min(c1, c2), r2: MAX_ROWS - 1, c2: Math.max(c1, c2), ra1: true, ca1: a.startsWith('$'), ra2: true, ca2: b.startsWith('$'), kind: 'cols' };
    }
    const [a, b] = txt.split(':');
    const r1 = Number(a.replace('$', '')) - 1;
    const r2 = Number(b.replace('$', '')) - 1;
    return { t: 'range', sheet: t.sheet, r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAX_COLS - 1, ra1: a.startsWith('$'), ca1: true, ra2: b.startsWith('$'), ca2: true, kind: 'rows' };
  };

  const primary = (): Node => {
    const t = next();
    if (!t) throw new FormulaError('Unexpected end of formula', src.length);
    switch (t.type) {
      case 'num':
        return { t: 'num', v: parseFloat(t.text) };
      case 'str':
        return { t: 'str', v: t.text.slice(1, -1).replace(/""/g, '"') };
      case 'bool':
        return { t: 'bool', v: t.text.toUpperCase() === 'TRUE' };
      case 'err':
        return { t: 'err', v: t.text.toUpperCase() };
      case 'ref': {
        let node = refNode(t);
        // A1:B2 (both sides may carry $, only the first a sheet)
        if (peek()?.type === 'colon' && toks[p + 1]?.type === 'ref' && !toks[p + 1].sheet) {
          p++;
          const b = next();
          const a = node as Extract<Node, { t: 'ref' }>;
          const bb = parseCellRef(b.refText!);
          const r1 = Math.min(a.r, bb.r);
          const r2 = Math.max(a.r, bb.r);
          const c1 = Math.min(a.c, bb.c);
          const c2 = Math.max(a.c, bb.c);
          node = {
            t: 'range',
            sheet: a.sheet,
            r1,
            c1,
            r2,
            c2,
            ra1: a.r <= bb.r ? a.ra : bb.ra,
            ca1: a.c <= bb.c ? a.ca : bb.ca,
            ra2: a.r <= bb.r ? bb.ra : a.ra,
            ca2: a.c <= bb.c ? bb.ca : a.ca,
            kind: 'cells',
          };
        }
        if (peek()?.type === 'hash') {
          p++;
          return { t: 'spill', ref: node };
        }
        return node;
      }
      case 'colrange':
      case 'rowrange':
        return refNode(t);
      case 'name':
        return { t: 'name', name: (t.refText ?? t.text).toUpperCase().replace(/^_XLPM\./, ''), sheet: t.sheet };
      case 'func': {
        const name = t.text.toUpperCase().replace(/^(_XLFN\.|_XLWS\.)+/, '');
        expect('lparen', '(');
        const args: Node[] = [];
        if (peek()?.type === 'rparen') {
          p++;
          return { t: 'func', name, args };
        }
        for (;;) {
          const tk = peek();
          if (tk && (tk.type === 'sep' || tk.type === 'semi')) {
            args.push({ t: 'missing' });
            p++;
            continue;
          }
          if (tk && tk.type === 'rparen') {
            args.push({ t: 'missing' });
            p++;
            break;
          }
          args.push(expr(0));
          const sep = next();
          if (!sep) throw new FormulaError('Missing closing parenthesis', src.length);
          if (sep.type === 'rparen') break;
          if (sep.type !== 'sep' && sep.type !== 'semi') throw new FormulaError(`Unexpected "${sep.text}"`, sep.start);
        }
        let node: Node = { t: 'func', name, args };
        // immediately-invoked LAMBDA: LAMBDA(x, x+1)(5)
        while (peek()?.type === 'lparen') {
          p++;
          const cargs: Node[] = [];
          if (peek()?.type === 'rparen') p++;
          else
            for (;;) {
              cargs.push(expr(0));
              const sep = next();
              if (!sep) throw new FormulaError('Missing closing parenthesis', src.length);
              if (sep.type === 'rparen') break;
            }
          node = { t: 'call', callee: node, args: cargs };
        }
        return node;
      }
      case 'lparen': {
        const e = expr(0);
        expect('rparen', ')');
        if (peek()?.type === 'hash') {
          p++;
          return { t: 'spill', ref: e };
        }
        return e;
      }
      case 'lbrace': {
        const rows: Node[][] = [[]];
        for (;;) {
          const tk = peek();
          if (!tk) throw new FormulaError('Missing }', src.length);
          if (tk.type === 'rbrace') {
            p++;
            break;
          }
          let sign = 1;
          if (tk.type === 'op' && (tk.text === '-' || tk.text === '+')) {
            sign = tk.text === '-' ? -1 : 1;
            p++;
          }
          const v = next();
          let node: Node;
          if (v.type === 'num') node = { t: 'num', v: sign * parseFloat(v.text) };
          else if (v.type === 'str') node = { t: 'str', v: v.text.slice(1, -1).replace(/""/g, '"') };
          else if (v.type === 'bool') node = { t: 'bool', v: v.text.toUpperCase() === 'TRUE' };
          else if (v.type === 'err') node = { t: 'err', v: v.text.toUpperCase() };
          else throw new FormulaError('Array constants may only contain values', v.start);
          rows[rows.length - 1].push(node);
          const sep = peek();
          if (sep?.type === 'sep') p++;
          else if (sep?.type === 'semi') {
            p++;
            rows.push([]);
          }
        }
        return { t: 'array', rows };
      }
      case 'op':
        if (t.text === '-' || t.text === '+' || t.text === '@') {
          const arg = expr(PREC_UNARY);
          return { t: 'unary', op: t.text, arg };
        }
        throw new FormulaError(`Unexpected "${t.text}"`, t.start);
      default:
        throw new FormulaError(`Unexpected "${t.text}"`, t.start);
    }
  };

  const expr = (minPrec: number): Node => {
    let left = primary();
    for (;;) {
      const t = peek();
      if (!t) break;
      if (t.type === 'op' && t.text === '%') {
        if (PREC_POSTFIX < minPrec) break;
        p++;
        left = { t: 'postfix', op: '%', arg: left };
        continue;
      }
      if (t.type === 'colon') {
        // range operator between expressions (e.g. INDEX(...):INDEX(...))
        p++;
        const right = primary();
        left = { t: 'bin', op: ':', left, right };
        continue;
      }
      if (t.type !== 'op') break;
      const prec = BIN_PREC[t.text];
      if (prec === undefined || prec < minPrec) break;
      p++;
      const right = expr(prec + 1);
      left = { t: 'bin', op: t.text, left, right };
    }
    return left;
  };

  if (!toks.length) throw new FormulaError('Empty formula', 0);
  const result = expr(0);
  if (p < toks.length) throw new FormulaError(`Unexpected "${toks[p].text}"`, toks[p].start);
  return result;
}

/* ============================================================ utilities */

/** Collects static references (cells, ranges, names) used by a formula. */
export function collectRefs(node: Node, out: Node[] = []): Node[] {
  switch (node.t) {
    case 'ref':
    case 'range':
    case 'name':
      out.push(node);
      break;
    case 'spill':
      out.push(node);
      collectRefs(node.ref, out);
      break;
    case 'func':
      node.args.forEach((a) => collectRefs(a, out));
      break;
    case 'call':
      collectRefs(node.callee, out);
      node.args.forEach((a) => collectRefs(a, out));
      break;
    case 'unary':
    case 'postfix':
      collectRefs(node.arg, out);
      break;
    case 'bin':
      collectRefs(node.left, out);
      collectRefs(node.right, out);
      break;
    default:
      break;
  }
  return out;
}

export function collectFunctions(node: Node, out = new Set<string>()): Set<string> {
  if (node.t === 'func') {
    out.add(node.name);
    node.args.forEach((a) => collectFunctions(a, out));
  } else if (node.t === 'call') {
    collectFunctions(node.callee, out);
    node.args.forEach((a) => collectFunctions(a, out));
  } else if (node.t === 'unary' || node.t === 'postfix') collectFunctions(node.arg, out);
  else if (node.t === 'bin') {
    collectFunctions(node.left, out);
    collectFunctions(node.right, out);
  } else if (node.t === 'spill') collectFunctions(node.ref, out);
  return out;
}
