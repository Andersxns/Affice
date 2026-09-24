import { tokenize, type Node, type Token } from './parser';
import { colIndex, colName, MAX_COLS, MAX_ROWS, quoteSheet, type Range } from '../model/address';

/* ======================================================== AST offset */

const REF_ERR: Node = { t: 'err', v: '#REF!' };

/** Shifts the relative references of a parsed formula (used for CF rules and fills). */
export function offsetAst(node: Node, dr: number, dc: number): Node {
  if (dr === 0 && dc === 0) return node;
  switch (node.t) {
    case 'ref': {
      const r = node.ra ? node.r : node.r + dr;
      const c = node.ca ? node.c : node.c + dc;
      if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return REF_ERR;
      return { ...node, r, c };
    }
    case 'range': {
      const r1 = node.kind === 'cols' || node.ra1 ? node.r1 : node.r1 + dr;
      const r2 = node.kind === 'cols' || node.ra2 ? node.r2 : node.r2 + dr;
      const c1 = node.kind === 'rows' || node.ca1 ? node.c1 : node.c1 + dc;
      const c2 = node.kind === 'rows' || node.ca2 ? node.c2 : node.c2 + dc;
      if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) return REF_ERR;
      return { ...node, r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
    }
    case 'spill':
      return { ...node, ref: offsetAst(node.ref, dr, dc) };
    case 'func':
      return { ...node, args: node.args.map((a) => offsetAst(a, dr, dc)) };
    case 'call':
      return { ...node, callee: offsetAst(node.callee, dr, dc), args: node.args.map((a) => offsetAst(a, dr, dc)) };
    case 'unary':
    case 'postfix':
      return { ...node, arg: offsetAst(node.arg, dr, dc) } as Node;
    case 'bin':
      return { ...node, left: offsetAst(node.left, dr, dc), right: offsetAst(node.right, dr, dc) };
    default:
      return node;
  }
}

/* ===================================================== text rewriting */

export interface RefInfo {
  sheet?: string;
  kind: 'cell' | 'range' | 'cols' | 'rows';
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  ra1: boolean;
  ca1: boolean;
  ra2: boolean;
  ca2: boolean;
}

function parseCell(text: string): { r: number; c: number; ra: boolean; ca: boolean } {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(text)!;
  return { ca: m[1] === '$', c: colIndex(m[2]), ra: m[3] === '$', r: Number(m[4]) - 1 };
}

function cellText(r: number, c: number, ra: boolean, ca: boolean): string {
  return `${ca ? '$' : ''}${colName(c)}${ra ? '$' : ''}${r + 1}`;
}

export function formatRef(info: RefInfo): string {
  const prefix = info.sheet !== undefined ? `${quoteSheet(info.sheet)}!` : '';
  switch (info.kind) {
    case 'cell':
      return prefix + cellText(info.r1, info.c1, info.ra1, info.ca1);
    case 'range':
      return prefix + cellText(info.r1, info.c1, info.ra1, info.ca1) + ':' + cellText(info.r2, info.c2, info.ra2, info.ca2);
    case 'cols':
      return `${prefix}${info.ca1 ? '$' : ''}${colName(info.c1)}:${info.ca2 ? '$' : ''}${colName(info.c2)}`;
    case 'rows':
      return `${prefix}${info.ra1 ? '$' : ''}${info.r1 + 1}:${info.ra2 ? '$' : ''}${info.r2 + 1}`;
  }
}

/**
 * Calls `fn` for every reference in a formula and splices in its replacement
 * (null keeps the original text). Whitespace and everything else is preserved.
 */
export function rewriteRefs(formula: string, fn: (ref: RefInfo, tok: Token, end: number) => string | null): string {
  const toks = tokenize(formula);
  const sig: number[] = [];
  toks.forEach((t, i) => {
    if (t.type !== 'ws') sig.push(i);
  });
  const edits: { start: number; end: number; text: string }[] = [];
  for (let s = 0; s < sig.length; s++) {
    const t = toks[sig[s]];
    let info: RefInfo | null = null;
    let end = t.end;
    if (t.type === 'ref') {
      const a = parseCell(t.refText!);
      const colon = toks[sig[s + 1]];
      const second = toks[sig[s + 2]];
      if (colon?.type === 'colon' && second?.type === 'ref' && second.sheet === undefined) {
        const b = parseCell(second.refText!);
        info = { sheet: t.sheet, kind: 'range', r1: a.r, c1: a.c, r2: b.r, c2: b.c, ra1: a.ra, ca1: a.ca, ra2: b.ra, ca2: b.ca };
        end = second.end;
        s += 2;
      } else info = { sheet: t.sheet, kind: 'cell', r1: a.r, c1: a.c, r2: a.r, c2: a.c, ra1: a.ra, ca1: a.ca, ra2: a.ra, ca2: a.ca };
    } else if (t.type === 'colrange') {
      const [p, q] = t.refText!.split(':');
      info = { sheet: t.sheet, kind: 'cols', r1: 0, r2: MAX_ROWS - 1, c1: colIndex(p.replace('$', '')), c2: colIndex(q.replace('$', '')), ra1: true, ra2: true, ca1: p.startsWith('$'), ca2: q.startsWith('$') };
    } else if (t.type === 'rowrange') {
      const [p, q] = t.refText!.split(':');
      info = { sheet: t.sheet, kind: 'rows', c1: 0, c2: MAX_COLS - 1, r1: Number(p.replace('$', '')) - 1, r2: Number(q.replace('$', '')) - 1, ca1: true, ca2: true, ra1: p.startsWith('$'), ra2: q.startsWith('$') };
    }
    if (!info) continue;
    const rep = fn(info, t, end);
    if (rep !== null) edits.push({ start: t.start, end, text: rep });
  }
  if (!edits.length) return formula;
  let out = '';
  let pos = 0;
  for (const e of edits) {
    out += formula.slice(pos, e.start) + e.text;
    pos = e.end;
  }
  return out + formula.slice(pos);
}

const sameSheet = (a: string | undefined, b: string) => (a ?? '').toLowerCase() === b.toLowerCase();

/* =========================================================== copy/fill */

/** Adjusts relative references for a formula copied by (dr, dc). */
export function translateFormula(formula: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return formula;
  return rewriteRefs(formula, (ref) => {
    const out = { ...ref };
    if (ref.kind !== 'cols') {
      if (!ref.ra1) out.r1 += dr;
      if (!ref.ra2) out.r2 += dr;
    }
    if (ref.kind !== 'rows') {
      if (!ref.ca1) out.c1 += dc;
      if (!ref.ca2) out.c2 += dc;
    }
    if (out.r1 < 0 || out.c1 < 0 || out.r2 < 0 || out.c2 < 0 || out.r1 >= MAX_ROWS || out.r2 >= MAX_ROWS || out.c1 >= MAX_COLS || out.c2 >= MAX_COLS) return withSheet(ref.sheet, '#REF!');
    if (out.kind === 'cell') {
      out.r2 = out.r1;
      out.c2 = out.c1;
    }
    return formatRef(out);
  });
}

function withSheet(sheet: string | undefined, text: string): string {
  return sheet !== undefined ? `${quoteSheet(sheet)}!${text}` : text;
}

/* ========================================================= structural */

export interface StructuralOp {
  sheet: string; // sheet name where rows/cols are inserted or deleted
  axis: 'row' | 'col';
  at: number;
  count: number; // > 0 insert, < 0 delete
}

/** Shifts a [lo, hi] interval for an insert/delete. Returns null if it vanishes. */
export function shiftSpan(lo: number, hi: number, at: number, count: number): [number, number] | null {
  if (count > 0) {
    return [lo >= at ? lo + count : lo, hi >= at ? hi + count : hi];
  }
  const n = -count;
  const end = at + n - 1;
  const nlo = lo < at ? lo : lo > end ? lo - n : at;
  const nhi = hi < at ? hi : hi > end ? hi - n : at - 1;
  if (nhi < nlo) return null;
  return [nlo, nhi];
}

export function adjustFormulaStructural(formula: string, formulaSheet: string, op: StructuralOp): string {
  return rewriteRefs(formula, (ref) => {
    if (!sameSheet(ref.sheet ?? formulaSheet, op.sheet)) return null;
    const out = { ...ref };
    const max = op.axis === 'row' ? MAX_ROWS : MAX_COLS;
    if (op.axis === 'row') {
      if (ref.kind === 'cols') return null;
      if (ref.kind === 'cell') {
        const s = shiftSpan(ref.r1, ref.r1, op.at, op.count);
        if (!s || s[0] >= max) return withSheet(ref.sheet, '#REF!');
        out.r1 = out.r2 = s[0];
      } else {
        const s = shiftSpan(ref.r1, ref.r2, op.at, op.count);
        if (!s) return withSheet(ref.sheet, '#REF!');
        out.r1 = s[0];
        out.r2 = Math.min(s[1], max - 1);
      }
    } else {
      if (ref.kind === 'rows') return null;
      if (ref.kind === 'cell') {
        const s = shiftSpan(ref.c1, ref.c1, op.at, op.count);
        if (!s || s[0] >= max) return withSheet(ref.sheet, '#REF!');
        out.c1 = out.c2 = s[0];
      } else {
        const s = shiftSpan(ref.c1, ref.c2, op.at, op.count);
        if (!s) return withSheet(ref.sheet, '#REF!');
        out.c1 = s[0];
        out.c2 = Math.min(s[1], max - 1);
      }
    }
    if (out.r1 === ref.r1 && out.r2 === ref.r2 && out.c1 === ref.c1 && out.c2 === ref.c2) return null;
    return formatRef(out);
  });
}

/** Shifts a model range (merges, CF, validation, charts) for an insert/delete. */
export function shiftRange(rg: Range, axis: 'row' | 'col', at: number, count: number): Range | null {
  if (axis === 'row') {
    const s = shiftSpan(rg.r1, rg.r2, at, count);
    return s ? { ...rg, r1: s[0], r2: Math.min(s[1], MAX_ROWS - 1) } : null;
  }
  const s = shiftSpan(rg.c1, rg.c2, at, count);
  return s ? { ...rg, c1: s[0], c2: Math.min(s[1], MAX_COLS - 1) } : null;
}

/* ================================================================ move */

export interface MoveOp {
  srcSheet: string;
  src: Range;
  dstSheet: string;
  dr: number;
  dc: number;
}

/** Re-points references into a moved block (cut & paste, drag-move). */
export function adjustFormulaMove(formula: string, formulaSheet: string, op: MoveOp): string {
  return rewriteRefs(formula, (ref) => {
    if (!sameSheet(ref.sheet ?? formulaSheet, op.srcSheet)) return null;
    if (ref.kind !== 'cell' && ref.kind !== 'range') return null;
    const { src } = op;
    const inside = ref.r1 >= src.r1 && ref.r2 <= src.r2 && ref.c1 >= src.c1 && ref.c2 <= src.c2;
    if (!inside) return null;
    const out = { ...ref, r1: ref.r1 + op.dr, r2: ref.r2 + op.dr, c1: ref.c1 + op.dc, c2: ref.c2 + op.dc };
    if (out.r1 < 0 || out.c1 < 0 || out.r2 >= MAX_ROWS || out.c2 >= MAX_COLS) return withSheet(ref.sheet, '#REF!');
    if (!sameSheet(op.dstSheet, op.srcSheet)) out.sheet = sameSheet(formulaSheet, op.dstSheet) ? undefined : op.dstSheet;
    return formatRef(out);
  });
}

/* ============================================================== sheets */

export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  const toks = tokenize(formula);
  let out = '';
  let pos = 0;
  for (const t of toks) {
    if (t.sheet !== undefined && sameSheet(t.sheet, oldName)) {
      out += formula.slice(pos, t.start) + `${quoteSheet(newName)}!${t.refText ?? ''}`;
      pos = t.end;
    }
  }
  return pos ? out + formula.slice(pos) : formula;
}

export function deleteSheetInFormula(formula: string, name: string): string {
  const toks = tokenize(formula);
  let out = '';
  let pos = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.sheet !== undefined && sameSheet(t.sheet, name)) {
      let end = t.end;
      // swallow the second half of a range (Sheet!A1:B2)
      const j = toks.findIndex((x, k) => k > i && x.type !== 'ws');
      if (t.type === 'ref' && j > 0 && toks[j].type === 'colon') {
        const k = toks.findIndex((x, q) => q > j && x.type !== 'ws');
        if (k > 0 && toks[k].type === 'ref' && toks[k].sheet === undefined) {
          end = toks[k].end;
          i = k;
        }
      }
      out += formula.slice(pos, t.start) + '#REF!';
      pos = end;
    }
  }
  return pos ? out + formula.slice(pos) : formula;
}

/** Does a formula reference any cell on the given sheet or inside a range? (used for copy decisions) */
export function formulaTouches(formula: string, formulaSheet: string, sheet: string, rg?: Range): boolean {
  let hit = false;
  rewriteRefs(formula, (ref) => {
    if (sameSheet(ref.sheet ?? formulaSheet, sheet) && (!rg || (ref.r1 <= rg.r2 && ref.r2 >= rg.r1 && ref.c1 <= rg.c2 && ref.c2 >= rg.c1))) hit = true;
    return null;
  });
  return hit;
}

/** Toggles $ markers of the reference under the caret (F4 in the formula editor). */
export function cycleAbsolute(formula: string, caret: number): { text: string; caret: number } {
  let result = { text: formula, caret };
  rewriteRefs(formula, (ref, tok, end) => {
    if (caret < tok.start || caret > end) return null;
    // $A$1 → A$1 → $A1 → A1 → $A$1
    const state = (ref.ra1 ? 2 : 0) + (ref.ca1 ? 1 : 0);
    const next = state === 3 ? 2 : state === 2 ? 1 : state === 1 ? 0 : 3;
    const out = { ...ref, ra1: next >= 2, ca1: next % 2 === 1, ra2: next >= 2, ca2: next % 2 === 1 };
    if (ref.kind === 'cols') out.ca1 = out.ca2 = !ref.ca1;
    else if (ref.kind === 'rows') out.ra1 = out.ra2 = !ref.ra1;
    const text = formatRef(out);
    result = { text: formula.slice(0, tok.start) + text + formula.slice(end), caret: tok.start + text.length };
    return null;
  });
  return result;
}
