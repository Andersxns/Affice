import { compare, ERR, isErr, asErr, isLambda, isRef, matrix, type RefValue, type Scalar, type Value } from '../values';
import { callLambda, clampRef, deref, numEquals, toGrid, type EvalContext } from '../evaluator';
import { addrName, MAX_COLS, MAX_ROWS, parseRange, quoteSheet } from '../../model/address';
import { define } from './registry';
import { bool, check, fn, fullDims, grid, int, num, optInt, str } from './helpers';

const LK = 'Lookup' as const;

/* ============================================================ matching */

function wildcardRe(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) re += '\\' + pattern[++i];
    else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'i');
}

/** Equality used by lookups: case-insensitive text, optional wildcards. */
function matcher(needle: Scalar, wildcards: boolean): (v: Scalar) => boolean {
  if (typeof needle === 'string') {
    if (wildcards && /[*?~]/.test(needle)) {
      const re = wildcardRe(needle);
      return (v) => typeof v === 'string' && re.test(v);
    }
    const low = needle.toLowerCase();
    return (v) => typeof v === 'string' && v.toLowerCase() === low;
  }
  if (typeof needle === 'number') return (v) => typeof v === 'number' && numEquals(v, needle);
  if (typeof needle === 'boolean') return (v) => v === needle;
  if (needle === null) return (v) => v === null || v === '';
  return () => false;
}

function sameType(a: Scalar, b: Scalar): boolean {
  return typeof a === typeof b;
}

/** Excel's approximate binary search: last position whose value is <= needle (ascending data). */
function bsearchAsc(values: Scalar[], needle: Scalar): number {
  let lo = 0;
  let hi = values.length - 1;
  let best = -1;
  while (lo <= hi) {
    let mid = (lo + hi) >> 1;
    // skip values of another type (Excel ignores them in approximate lookups)
    let m = mid;
    while (m >= lo && !sameType(values[m], needle)) m--;
    if (m < lo) {
      lo = mid + 1;
      continue;
    }
    mid = m;
    const c = compare(values[mid], needle);
    if (c <= 0) {
      best = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return best;
}

function bsearchDesc(values: Scalar[], needle: Scalar): number {
  // MATCH type -1: smallest value >= needle, data descending
  let best = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (!sameType(v, needle)) continue;
    if (compare(v, needle) >= 0) best = i;
    else break;
  }
  return best;
}

/** A one-dimensional vector of values from a reference or array (row or column). */
function vector(v: Value, ctx: EvalContext): Scalar[] {
  const g = toGrid(v, ctx);
  if (g.length === 1) return g[0];
  if ((g[0]?.length ?? 0) === 1) return g.map((r) => r[0]);
  throw ERR.NA;
}

function needleOf(v: Scalar): Scalar {
  if (isErr(v)) throw asErr(v);
  return v === null ? 0 : v;
}

/* ============================================================ VLOOKUP */

function tableLookup(needle: Scalar, table: Value, index: number, approx: boolean, ctx: EvalContext, horizontal: boolean): Scalar {
  const g = toGrid(table, ctx);
  const rows = g.length;
  const cols = g[0]?.length ?? 0;
  const n = horizontal ? cols : rows;
  const width = horizontal ? rows : cols;
  if (index < 1) throw ERR.VALUE;
  if (index > width) throw ERR.REF;
  const key = (i: number) => (horizontal ? g[0][i] : g[i][0]);
  const pick = (i: number) => {
    const v = horizontal ? g[index - 1][i] : g[i][index - 1];
    return v === null ? 0 : v;
  };
  const x = needleOf(needle);
  if (!approx) {
    const test = matcher(x, true);
    for (let i = 0; i < n; i++) if (test(key(i))) return pick(i);
    throw ERR.NA;
  }
  const keys: Scalar[] = new Array(n);
  for (let i = 0; i < n; i++) keys[i] = key(i);
  const i = bsearchAsc(keys, x);
  if (i < 0) throw ERR.NA;
  return pick(i);
}

fn('VLOOKUP', LK, 3, 4, 'vrvv', 'lookup_value, table_array, col_index_num, [range_lookup]', 'Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify.', (args: Value[], ctx) =>
  tableLookup(args[0] as Scalar, args[1], int(args[2] as Scalar), args.length > 3 ? bool(args[3] as Scalar) : true, ctx, false),
);
fn('HLOOKUP', LK, 3, 4, 'vrvv', 'lookup_value, table_array, row_index_num, [range_lookup]', 'Looks for a value in the top row of a table and returns the value in the same column from a row you specify.', (args: Value[], ctx) =>
  tableLookup(args[0] as Scalar, args[1], int(args[2] as Scalar), args.length > 3 ? bool(args[3] as Scalar) : true, ctx, true),
);

fn('LOOKUP', LK, 2, 3, 'vrr', 'lookup_value, lookup_vector, [result_vector]', 'Looks up a value either from a one-row or one-column range or from an array.', (args: Value[], ctx) => {
  const x = needleOf(args[0] as Scalar);
  const g = toGrid(args[1], ctx);
  let keys: Scalar[];
  let results: Scalar[];
  if (args.length > 2) {
    keys = vector(args[1], ctx);
    results = vector(args[2], ctx);
  } else {
    const rows = g.length;
    const cols = g[0]?.length ?? 0;
    if (cols > rows) {
      keys = g[0];
      results = g[rows - 1];
    } else {
      keys = g.map((r) => r[0]);
      results = g.map((r) => r[cols - 1]);
    }
  }
  const i = bsearchAsc(keys, x);
  if (i < 0) throw ERR.NA;
  const v = results[i];
  if (v === undefined) throw ERR.NA;
  return v === null ? 0 : v;
});

/* ============================================================== MATCH */

function matchPos(x: Scalar, values: Scalar[], type: number): number {
  if (type === 0) {
    const test = matcher(x, true);
    for (let i = 0; i < values.length; i++) if (test(values[i])) return i;
    return -1;
  }
  return type > 0 ? bsearchAsc(values, x) : bsearchDesc(values, x);
}

fn('MATCH', LK, 2, 3, 'vrv', 'lookup_value, lookup_array, [match_type]', 'Returns the relative position of an item in an array that matches a specified value.', (args: Value[], ctx) => {
  const x = needleOf(args[0] as Scalar);
  const type = args.length > 2 ? Math.sign(num(args[2] as Scalar)) : 1;
  const i = matchPos(x, vector(args[1], ctx), type);
  if (i < 0) throw ERR.NA;
  return i + 1;
});

/**
 * XMATCH / XLOOKUP search. match_mode: 0 exact, -1 exact or next smaller, 1 exact or next larger,
 * 2 wildcard; search_mode: 1 first→last, -1 last→first, 2 binary ascending, -2 binary descending.
 */
function xfind(x: Scalar, values: Scalar[], matchMode: number, searchMode: number): number {
  const n = values.length;
  if (searchMode === 2 || searchMode === -2) {
    // binary search
    let lo = 0;
    let hi = n - 1;
    let found = -1;
    let below = -1;
    let above = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      let c = compare(values[mid], x);
      if (searchMode === -2) c = -c;
      if (c === 0 && sameType(values[mid], x)) {
        found = mid;
        break;
      }
      if (c < 0) {
        if (searchMode === 2) below = mid;
        else above = mid;
        lo = mid + 1;
      } else {
        if (searchMode === 2) above = mid;
        else below = mid;
        hi = mid - 1;
      }
    }
    if (found >= 0) return found;
    if (matchMode === -1) return below;
    if (matchMode === 1) return above;
    return -1;
  }
  const order = searchMode === -1 ? [...Array(n).keys()].reverse() : [...Array(n).keys()];
  const test = matcher(x, matchMode === 2);
  for (const i of order) if (test(values[i])) return i;
  if (matchMode === -1 || matchMode === 1) {
    let best = -1;
    for (const i of order) {
      const v = values[i];
      if (!sameType(v, x)) continue;
      const c = compare(v, x);
      if (matchMode === -1 && c < 0 && (best < 0 || compare(v, values[best]) > 0)) best = i;
      if (matchMode === 1 && c > 0 && (best < 0 || compare(v, values[best]) < 0)) best = i;
    }
    return best;
  }
  return -1;
}

fn('XMATCH', LK, 2, 4, 'vrvv', 'lookup_value, lookup_array, [match_mode], [search_mode]', 'Returns the relative position of an item in an array (exact match by default).', (args: Value[], ctx) => {
  const x = needleOf(args[0] as Scalar);
  const mm = args.length > 2 && args[2] !== null ? int(args[2] as Scalar) : 0;
  const sm = args.length > 3 && args[3] !== null ? int(args[3] as Scalar) : 1;
  check([0, -1, 1, 2].includes(mm) && [1, -1, 2, -2].includes(sm), ERR.VALUE);
  const i = xfind(x, vector(args[1], ctx), mm, sm);
  if (i < 0) throw ERR.NA;
  return i + 1;
});

define({
  name: 'XLOOKUP',
  category: LK,
  min: 3,
  max: 6,
  args: 'vrrevv',
  syntax: 'lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]',
  desc: 'Searches a range or an array for a match and returns the corresponding item from a second range or array.',
  fn: (args: Value[], ctx: EvalContext) => {
    const x = needleOf(args[0] as Scalar);
    const look = args[1];
    const ret = args[2];
    const mm = args.length > 4 && args[4] !== null ? int(args[4] as Scalar) : 0;
    const sm = args.length > 5 && args[5] !== null ? int(args[5] as Scalar) : 1;
    check([0, -1, 1, 2].includes(mm) && [1, -1, 2, -2].includes(sm), ERR.VALUE);
    const [lr, lc] = fullDims(look);
    const vertical = lc === 1;
    if (!vertical && lr !== 1) throw ERR.VALUE;
    const i = xfind(x, vector(look, ctx), mm, sm);
    if (i < 0) {
      if (args.length > 3) return args[3] as Scalar;
      throw ERR.NA;
    }
    const [rr, rc] = fullDims(ret);
    if (vertical && rr !== lr) throw ERR.VALUE;
    if (!vertical && rc !== lc) throw ERR.VALUE;
    if (isRef(ret)) {
      return vertical ? { ...ret, r1: ret.r1 + i, r2: ret.r1 + i } : { ...ret, c1: ret.c1 + i, c2: ret.c1 + i };
    }
    const g = toGrid(ret, ctx);
    if (vertical) return g[i].length === 1 ? g[i][0] : matrix([g[i]]);
    const col = g.map((row) => [row[i]]);
    return col.length === 1 ? col[0][0] : matrix(col);
  },
});

/* ============================================================== INDEX */

define({
  name: 'INDEX',
  category: LK,
  min: 1,
  max: 4,
  args: 'rvvv',
  syntax: 'array, [row_num], [column_num], [area_num]',
  desc: 'Returns a value or reference of the cell at the intersection of a particular row and column, in a given range.',
  fn: (args: Value[], ctx: EvalContext) => {
    const a = args[0];
    let r = args.length > 1 && args[1] !== null ? int(args[1] as Scalar) : 0;
    let c = args.length > 2 && args[2] !== null ? int(args[2] as Scalar) : 0;
    const [R, C] = fullDims(a);
    if (args.length === 2 && R === 1 && C > 1) {
      // INDEX(row_vector, n) picks the n-th column
      c = r;
      r = 1;
    } else if (args.length === 2 && C === 1) c = 1;
    if (r < 0 || c < 0 || r > R || c > C) throw ERR.REF;
    if (isRef(a)) {
      return {
        kind: 'ref',
        sheet: a.sheet,
        r1: r ? a.r1 + r - 1 : a.r1,
        r2: r ? a.r1 + r - 1 : a.r2,
        c1: c ? a.c1 + c - 1 : a.c1,
        c2: c ? a.c1 + c - 1 : a.c2,
      } satisfies RefValue;
    }
    const g = toGrid(a, ctx);
    if (r && c) return g[r - 1][c - 1];
    if (r) return matrix([g[r - 1]]);
    if (c) return matrix(g.map((row) => [row[c - 1]]));
    return matrix(g);
  },
});

/* ============================================================ OFFSET */

define({
  name: 'OFFSET',
  category: LK,
  min: 3,
  max: 5,
  args: 'rvvvv',
  volatile: true,
  syntax: 'reference, rows, cols, [height], [width]',
  desc: 'Returns a reference offset from a given reference by a number of rows and columns.',
  fn: (args: Value[]) => {
    const ref = args[0];
    if (!isRef(ref)) throw ERR.VALUE;
    const dr = int(args[1] as Scalar);
    const dc = int(args[2] as Scalar);
    const h = args.length > 3 && args[3] !== null ? int(args[3] as Scalar) : ref.r2 - ref.r1 + 1;
    const w = args.length > 4 && args[4] !== null ? int(args[4] as Scalar) : ref.c2 - ref.c1 + 1;
    if (h === 0 || w === 0) throw ERR.REF;
    let r1 = ref.r1 + dr;
    let c1 = ref.c1 + dc;
    let r2 = r1 + (h > 0 ? h - 1 : h + 1);
    let c2 = c1 + (w > 0 ? w - 1 : w + 1);
    [r1, r2] = [Math.min(r1, r2), Math.max(r1, r2)];
    [c1, c2] = [Math.min(c1, c2), Math.max(c1, c2)];
    if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) throw ERR.REF;
    return { kind: 'ref', sheet: ref.sheet, r1, c1, r2, c2 };
  },
});

/* ========================================================== INDIRECT */

function r1c1ToRef(text: string, ctx: EvalContext): { r1: number; c1: number; r2: number; c2: number } | null {
  const part = (s: string) => {
    const m = /^R(\[?-?\d*\]?)C(\[?-?\d*\]?)$/i.exec(s);
    if (!m) return null;
    const conv = (x: string, base: number) => {
      if (x === '') return base;
      if (x.startsWith('[')) return base + Number(x.slice(1, -1));
      return Number(x) - 1;
    };
    return { r: conv(m[1], ctx.row), c: conv(m[2], ctx.col) };
  };
  const [a, b] = text.split(':');
  const p = part(a);
  if (!p) return null;
  const q = b ? part(b) : p;
  if (!q) return null;
  return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) };
}

define({
  name: 'INDIRECT',
  category: LK,
  min: 1,
  max: 2,
  args: 'v',
  volatile: true,
  syntax: 'ref_text, [a1]',
  desc: 'Returns the reference specified by a text string.',
  fn: (args: Scalar[], ctx: EvalContext) => {
    let text = str(args[0]).trim();
    const a1 = args.length > 1 ? bool(args[1]) : true;
    let sheet = ctx.sheet;
    const bang = text.lastIndexOf('!');
    if (bang > 0) {
      let name = text.slice(0, bang);
      if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
      const id = ctx.sheetId(name);
      if (id === undefined) throw ERR.REF;
      sheet = id;
      text = text.slice(bang + 1);
    }
    const rg = a1 ? parseRange(text) : r1c1ToRef(text, ctx);
    if (!rg) {
      // defined name
      const dn = ctx.definedName(text.toUpperCase(), sheet);
      if (dn && (dn.t === 'ref' || dn.t === 'range')) {
        const sid = dn.sheet ? ctx.sheetId(dn.sheet) : sheet;
        if (sid === undefined) throw ERR.REF;
        return dn.t === 'ref' ? { kind: 'ref', sheet: sid, r1: dn.r, c1: dn.c, r2: dn.r, c2: dn.c } : { kind: 'ref', sheet: sid, r1: dn.r1, c1: dn.c1, r2: dn.r2, c2: dn.c2 };
      }
      throw ERR.REF;
    }
    if (rg.r1 < 0 || rg.c1 < 0) throw ERR.REF;
    return { kind: 'ref', sheet, ...rg };
  },
});

/* ============================================= ROW / COLUMN / ADDRESS */

define({
  name: 'ROW',
  category: LK,
  min: 0,
  max: 1,
  args: 'r',
  syntax: '[reference]',
  desc: 'Returns the row number of a reference.',
  fn: (args: Value[], ctx: EvalContext) => {
    if (!args.length || args[0] === null) return ctx.row + 1;
    const ref = args[0];
    if (!isRef(ref)) throw ERR.VALUE;
    if (ref.r1 === ref.r2) return ref.r1 + 1;
    const rg = clampRef(ref, ctx);
    return matrix(Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => [rg.r1 + i + 1]));
  },
});

define({
  name: 'COLUMN',
  category: LK,
  min: 0,
  max: 1,
  args: 'r',
  syntax: '[reference]',
  desc: 'Returns the column number of a reference.',
  fn: (args: Value[], ctx: EvalContext) => {
    if (!args.length || args[0] === null) return ctx.col + 1;
    const ref = args[0];
    if (!isRef(ref)) throw ERR.VALUE;
    if (ref.c1 === ref.c2) return ref.c1 + 1;
    const rg = clampRef(ref, ctx);
    return matrix([Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i + 1)]);
  },
});

fn('ROWS', LK, 1, 1, 'r', 'array', 'Returns the number of rows in a reference or array.', ([a]: Value[]) => {
  if (isLambda(a)) throw ERR.VALUE;
  return fullDims(a)[0];
});
fn('COLUMNS', LK, 1, 1, 'r', 'array', 'Returns the number of columns in a reference or array.', ([a]: Value[]) => {
  if (isLambda(a)) throw ERR.VALUE;
  return fullDims(a)[1];
});
fn('AREAS', LK, 1, 1, 'r', 'reference', 'Returns the number of areas in a reference.', ([a]: Value[]) => {
  if (!isRef(a)) throw ERR.VALUE;
  return 1;
});

fn('ADDRESS', LK, 2, 5, 'v', 'row_num, column_num, [abs_num], [a1], [sheet_text]', 'Creates a cell reference as text, given row and column numbers.', (args: Scalar[]) => {
  const r = int(args[0]);
  const c = int(args[1]);
  const abs = optInt(args, 2, 1);
  const a1 = args.length > 3 && args[3] !== null ? bool(args[3]) : true;
  const sheet = args.length > 4 && args[4] !== null ? str(args[4]) : '';
  check(r >= 1 && r <= MAX_ROWS && c >= 1 && c <= MAX_COLS && abs >= 1 && abs <= 4, ERR.VALUE);
  const absR = abs === 1 || abs === 2;
  const absC = abs === 1 || abs === 3;
  let out: string;
  if (a1) out = addrName(r - 1, c - 1, absR, absC);
  else out = `R${absR ? r : `[${r}]`}C${absC ? c : `[${c}]`}`;
  return sheet ? `${quoteSheet(sheet)}!${out}` : out;
});

/* ============================================ FORMULATEXT / HYPERLINK */

fn('FORMULATEXT', LK, 1, 1, 'r', 'reference', 'Returns the formula at the given reference as text.', ([a]: Value[], ctx) => {
  if (!isRef(a)) throw ERR.NA;
  const f = ctx.formulaText(a.sheet, a.r1, a.c1);
  if (f === undefined) throw ERR.NA;
  return '=' + f;
});

fn('HYPERLINK', LK, 1, 2, 'v', 'link_location, [friendly_name]', 'Creates a shortcut that opens a web page or jumps to a location (Ctrl+click in Affice).', (args: Scalar[]) => {
  if (args.length > 1 && args[1] !== null) {
    const f = args[1];
    if (isErr(f)) throw asErr(f);
    return f;
  }
  return str(args[0]);
});

fn('TRANSPOSE', LK, 1, 1, 'r', 'array', 'Converts a vertical range to a horizontal range, or vice versa.', ([a]: Value[], ctx) => {
  const g = toGrid(a, ctx);
  const R = g.length;
  const C = g[0]?.length ?? 0;
  const out: Scalar[][] = [];
  for (let c = 0; c < C; c++) {
    const row: Scalar[] = [];
    for (let r = 0; r < R; r++) row.push(g[r][c] === null ? 0 : g[r][c]);
    out.push(row);
  }
  return matrix(out);
});

/* ======================================================= dynamic arrays */

function sortKeyCompare(a: Scalar, b: Scalar): number {
  // blanks sort last, errors after everything
  const ea = isErr(a);
  const eb = isErr(b);
  if (ea || eb) return ea && eb ? 0 : ea ? 1 : -1;
  if (a === null || b === null) return a === b ? 0 : a === null ? 1 : -1;
  return compare(a, b);
}

fn('SORT', LK, 1, 4, 'rvvv', 'array, [sort_index], [sort_order], [by_col]', 'Sorts the contents of a range or array.', (args: Value[], ctx) => {
  let g = grid(args[0], ctx).map((r) => r.slice());
  const byCol = args.length > 3 && args[3] !== null ? bool(args[3] as Scalar) : false;
  if (byCol) g = transpose(g);
  const idx = args.length > 1 && args[1] !== null ? int(args[1] as Scalar) : 1;
  const order = args.length > 2 && args[2] !== null ? int(args[2] as Scalar) : 1;
  check(idx >= 1 && idx <= (g[0]?.length ?? 0), ERR.VALUE);
  check(order === 1 || order === -1, ERR.VALUE);
  const sorted = g
    .map((row, i) => ({ row, i }))
    .sort((p, q) => {
      const a = p.row[idx - 1];
      const b = q.row[idx - 1];
      let c = sortKeyCompare(a, b);
      // blanks and errors stay last in both directions
      if (order === -1 && a !== null && b !== null && !isErr(a) && !isErr(b)) c = -c;
      return c || p.i - q.i;
    })
    .map((x) => x.row.map((v) => (v === null ? 0 : v)));
  return matrix(byCol ? transpose(sorted) : sorted);
});

fn('SORTBY', LK, 2, 253, 'r', 'array, by_array1, [sort_order1], [by_array2, sort_order2], …', 'Sorts the contents of a range or array based on the values in a corresponding range or array.', (args: Value[], ctx) => {
  const g = grid(args[0], ctx);
  const keys: { vals: Scalar[]; order: number }[] = [];
  for (let i = 1; i < args.length; i += 2) {
    const vals = vector(args[i], ctx);
    const order = i + 1 < args.length && args[i + 1] !== null ? int(deref(args[i + 1], ctx) as Scalar) : 1;
    check(order === 1 || order === -1, ERR.VALUE);
    keys.push({ vals, order });
  }
  const R = g.length;
  const C = g[0]?.length ?? 0;
  const byCol = keys[0].vals.length === C && keys[0].vals.length !== R;
  const n = byCol ? C : R;
  for (const k of keys) if (k.vals.length !== n) throw ERR.VALUE;
  const idx = [...Array(n).keys()].sort((a, b) => {
    for (const k of keys) {
      const c = sortKeyCompare(k.vals[a], k.vals[b]);
      if (c) return k.order === 1 ? c : -c;
    }
    return a - b;
  });
  const fix = (v: Scalar) => (v === null ? 0 : v);
  if (byCol) return matrix(g.map((row) => idx.map((j) => fix(row[j]))));
  return matrix(idx.map((i) => g[i].map(fix)));
});

function transpose(g: Scalar[][]): Scalar[][] {
  const R = g.length;
  const C = g[0]?.length ?? 0;
  const out: Scalar[][] = [];
  for (let c = 0; c < C; c++) {
    const row: Scalar[] = [];
    for (let r = 0; r < R; r++) row.push(g[r][c]);
    out.push(row);
  }
  return out;
}

define({
  name: 'FILTER',
  category: LK,
  min: 2,
  max: 3,
  args: 'rre',
  syntax: 'array, include, [if_empty]',
  desc: 'Filters a range or array based on criteria you define.',
  fn: (args: Value[], ctx: EvalContext) => {
    const g = grid(args[0], ctx);
    const inc = grid(args[1], ctx);
    const R = g.length;
    const C = g[0]?.length ?? 0;
    const truth = (v: Scalar) => {
      if (isErr(v)) throw asErr(v);
      if (typeof v === 'number') return v !== 0;
      if (typeof v === 'boolean') return v;
      if (v === null || v === '') return false;
      throw ERR.VALUE;
    };
    let out: Scalar[][];
    if (inc.length === R && (inc[0]?.length ?? 0) === 1) out = g.filter((_, i) => truth(inc[i][0]));
    else if (inc.length === 1 && inc[0].length === C) {
      const keep = inc[0].map(truth);
      out = g.map((row) => row.filter((_, j) => keep[j]));
      if (!out[0]?.length) out = [];
    } else throw ERR.VALUE;
    if (!out.length) {
      if (args.length > 2) return args[2] as Scalar;
      throw ERR.CALC;
    }
    return matrix(out.map((row) => row.map((v) => (v === null ? 0 : v))));
  },
});

fn('UNIQUE', LK, 1, 3, 'rvv', 'array, [by_col], [exactly_once]', 'Returns the unique values from a range or array.', (args: Value[], ctx) => {
  let g = grid(args[0], ctx);
  const byCol = args.length > 1 && args[1] !== null ? bool(args[1] as Scalar) : false;
  const once = args.length > 2 && args[2] !== null ? bool(args[2] as Scalar) : false;
  if (byCol) g = transpose(g);
  const keyOf = (row: Scalar[]) => row.map((v) => (typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + String(isErr(v) ? asErr(v).error : v))).join('\u0001');
  const counts = new Map<string, number>();
  const firsts = new Map<string, Scalar[]>();
  for (const row of g) {
    const k = keyOf(row);
    counts.set(k, (counts.get(k) ?? 0) + 1);
    if (!firsts.has(k)) firsts.set(k, row);
  }
  const out = [...firsts.entries()].filter(([k]) => !once || counts.get(k) === 1).map(([, row]) => row.map((v) => (v === null ? 0 : v)));
  if (!out.length) throw ERR.CALC;
  return matrix(byCol ? transpose(out) : out);
});

fn('TAKE', LK, 2, 3, 'rvv', 'array, rows, [columns]', 'Returns a specified number of contiguous rows or columns from the start or end of an array.', (args: Value[], ctx) => {
  const g = grid(args[0], ctx);
  const rows = args[1] !== null ? int(args[1] as Scalar) : g.length;
  const cols = args.length > 2 && args[2] !== null ? int(args[2] as Scalar) : g[0].length;
  if (rows === 0 || cols === 0) throw ERR.CALC;
  const rs = rows > 0 ? g.slice(0, rows) : g.slice(Math.max(0, g.length + rows));
  return matrix(rs.map((row) => (cols > 0 ? row.slice(0, cols) : row.slice(Math.max(0, row.length + cols)))));
});

fn('DROP', LK, 2, 3, 'rvv', 'array, rows, [columns]', 'Excludes a specified number of rows or columns from the start or end of an array.', (args: Value[], ctx) => {
  const g = grid(args[0], ctx);
  const rows = args[1] !== null ? int(args[1] as Scalar) : 0;
  const cols = args.length > 2 && args[2] !== null ? int(args[2] as Scalar) : 0;
  const rs = rows >= 0 ? g.slice(rows) : g.slice(0, g.length + rows);
  const out = rs.map((row) => (cols >= 0 ? row.slice(cols) : row.slice(0, row.length + cols)));
  if (!out.length || !out[0].length) throw ERR.CALC;
  return matrix(out);
});

fn('CHOOSEROWS', LK, 2, 255, 'rv', 'array, row_num1, [row_num2], …', 'Returns the specified rows from an array.', (args: Value[], ctx) => {
  const g = grid(args[0], ctx);
  const idx: number[] = [];
  for (const a of args.slice(1)) for (const v of toGrid(a, ctx).flat()) idx.push(int(v));
  return matrix(
    idx.map((i) => {
      const k = i < 0 ? g.length + i : i - 1;
      if (i === 0 || k < 0 || k >= g.length) throw ERR.VALUE;
      return g[k];
    }),
  );
});

fn('CHOOSECOLS', LK, 2, 255, 'rv', 'array, col_num1, [col_num2], …', 'Returns the specified columns from an array.', (args: Value[], ctx) => {
  const g = grid(args[0], ctx);
  const C = g[0]?.length ?? 0;
  const idx: number[] = [];
  for (const a of args.slice(1)) for (const v of toGrid(a, ctx).flat()) idx.push(int(v));
  const ks = idx.map((i) => {
    const k = i < 0 ? C + i : i - 1;
    if (i === 0 || k < 0 || k >= C) throw ERR.VALUE;
    return k;
  });
  return matrix(g.map((row) => ks.map((k) => row[k])));
});

fn('VSTACK', LK, 1, 254, 'r', 'array1, [array2], …', 'Appends arrays vertically and in sequence to return a larger array.', (args: Value[], ctx) => {
  const gs = args.map((a) => grid(a, ctx));
  const width = Math.max(...gs.map((g) => g[0]?.length ?? 0));
  const out: Scalar[][] = [];
  for (const g of gs) for (const row of g) out.push([...row.map((v) => (v === null ? 0 : v)), ...Array(width - row.length).fill(ERR.NA)]);
  return matrix(out);
});

fn('HSTACK', LK, 1, 254, 'r', 'array1, [array2], …', 'Appends arrays horizontally and in sequence to return a larger array.', (args: Value[], ctx) => {
  const gs = args.map((a) => grid(a, ctx));
  const height = Math.max(...gs.map((g) => g.length));
  const out: Scalar[][] = [];
  for (let r = 0; r < height; r++) {
    const row: Scalar[] = [];
    for (const g of gs) {
      const w = g[0]?.length ?? 0;
      for (let c = 0; c < w; c++) {
        const v = g[r]?.[c];
        row.push(v === undefined ? ERR.NA : v === null ? 0 : v);
      }
    }
    out.push(row);
  }
  return matrix(out);
});

function toVector(args: Value[], ctx: EvalContext, asRow: boolean): Value {
  const g = grid(args[0], ctx);
  const ignore = args.length > 1 && args[1] !== null ? int(args[1] as Scalar) : 0;
  const byCol = args.length > 2 && args[2] !== null ? bool(args[2] as Scalar) : false;
  const src = byCol ? transpose(g) : g;
  const vals: Scalar[] = [];
  for (const row of src)
    for (const v of row) {
      if ((ignore === 1 || ignore === 3) && v === null) continue;
      if ((ignore === 2 || ignore === 3) && isErr(v)) continue;
      vals.push(v === null ? 0 : v);
    }
  if (!vals.length) throw ERR.CALC;
  return matrix(asRow ? [vals] : vals.map((v) => [v]));
}

fn('TOCOL', LK, 1, 3, 'rvv', 'array, [ignore], [scan_by_column]', 'Returns the array in a single column.', (args, ctx) => toVector(args, ctx, false));
fn('TOROW', LK, 1, 3, 'rvv', 'array, [ignore], [scan_by_column]', 'Returns the array in a single row.', (args, ctx) => toVector(args, ctx, true));

function wrap(args: Value[], ctx: EvalContext, byRows: boolean): Value {
  const vals = toGrid(args[0], ctx).flat();
  const n = int(deref(args[1], ctx) as Scalar);
  if (n < 1) throw ERR.NUM;
  const pad: Scalar = args.length > 2 ? (deref(args[2], ctx) as Scalar) : ERR.NA;
  const chunks: Scalar[][] = [];
  for (let i = 0; i < vals.length; i += n) {
    const chunk: Scalar[] = vals.slice(i, i + n).map((v) => (v === null ? 0 : v));
    while (chunk.length < n) chunk.push(pad);
    chunks.push(chunk);
  }
  return matrix(byRows ? chunks : transpose(chunks));
}

fn('WRAPROWS', LK, 2, 3, 'r', 'vector, wrap_count, [pad_with]', 'Wraps the provided row or column of values by rows after a specified number of elements.', (args, ctx) => wrap(args, ctx, true));
fn('WRAPCOLS', LK, 2, 3, 'r', 'vector, wrap_count, [pad_with]', 'Wraps the provided row or column of values by columns after a specified number of elements.', (args, ctx) => wrap(args, ctx, false));

fn('EXPAND', LK, 2, 4, 'rvve', 'array, rows, [columns], [pad_with]', 'Expands or pads an array to specified row and column dimensions.', (args: Value[], ctx) => {
  const g = grid(args[0], ctx);
  const rows = args[1] !== null ? int(args[1] as Scalar) : g.length;
  const cols = args.length > 2 && args[2] !== null ? int(args[2] as Scalar) : g[0].length;
  const pad: Scalar = args.length > 3 ? (args[3] as Scalar) : ERR.NA;
  if (rows < g.length || cols < (g[0]?.length ?? 0)) throw ERR.VALUE;
  const out: Scalar[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) {
      const v = g[r]?.[c];
      row.push(v === undefined ? pad : v === null ? 0 : v);
    }
    out.push(row);
  }
  return matrix(out);
});

/* ================================================== GROUPBY / PIVOTBY */

function aggregateLambda(f: Value, values: Scalar[], ctx: EvalContext): Scalar {
  if (!isLambda(f)) throw ERR.VALUE;
  const res = callLambda(f, [matrix(values.map((v) => [v]))], ctx);
  const g = toGrid(res, ctx);
  return g[0]?.[0] ?? 0;
}

fn('GROUPBY', LK, 3, 8, 'rrrvvvrv', 'row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array], [field_relationship]', 'Groups values by row fields and aggregates them (e.g. with SUM).', (args: Value[], ctx) => {
  const keys = grid(args[0], ctx);
  const vals = grid(args[1], ctx);
  if (keys.length !== vals.length) throw ERR.VALUE;
  const f = args[2];
  const totals = args.length > 4 && args[4] !== null ? int(args[4] as Scalar) : 1;
  const sortOrder = args.length > 5 && args[5] !== null ? int(args[5] as Scalar) : 1;
  const filter = args.length > 6 && args[6] !== null ? grid(args[6], ctx) : null;
  const groups = new Map<string, { key: Scalar[]; rows: number[] }>();
  keys.forEach((k, i) => {
    if (filter && !filter[i]?.[0]) return;
    const id = k.map((v) => (typeof v === 'string' ? v.toLowerCase() : String(v))).join('\u0001');
    if (!groups.has(id)) groups.set(id, { key: k, rows: [] });
    groups.get(id)!.rows.push(i);
  });
  const width = vals[0]?.length ?? 1;
  let out: Scalar[][] = [...groups.values()].map((gr) => [
    ...gr.key.map((v) => (v === null ? 0 : v)),
    ...Array.from({ length: width }, (_, c) => aggregateLambda(f, gr.rows.map((r) => vals[r][c]), ctx)),
  ]);
  const col = Math.abs(sortOrder) - 1;
  if (col >= 0 && col < (out[0]?.length ?? 0)) out = out.sort((p, q) => (sortOrder > 0 ? 1 : -1) * sortKeyCompare(p[col], q[col]));
  if (totals !== 0) {
    const all = keys.map((_, i) => i).filter((i) => !filter || filter[i]?.[0]);
    out.push(['Total', ...Array(Math.max(0, (keys[0]?.length ?? 1) - 1)).fill(''), ...Array.from({ length: width }, (_, c) => aggregateLambda(f, all.map((r) => vals[r][c]), ctx))]);
  }
  if (!out.length) throw ERR.CALC;
  return matrix(out);
});

fn('PIVOTBY', LK, 4, 11, 'rrrr', 'row_fields, col_fields, values, function, …', 'Groups values by row and column fields and aggregates them.', (args: Value[], ctx) => {
  const rk = grid(args[0], ctx).map((r) => r[0]);
  const ck = grid(args[1], ctx).map((r) => r[0]);
  const vals = grid(args[2], ctx).map((r) => r[0]);
  const f = args[3];
  if (rk.length !== ck.length || rk.length !== vals.length) throw ERR.VALUE;
  const norm = (v: Scalar) => (typeof v === 'string' ? v.toLowerCase() : String(v));
  const rows: Scalar[] = [];
  const cols: Scalar[] = [];
  for (const v of rk) if (!rows.some((x) => norm(x) === norm(v))) rows.push(v);
  for (const v of ck) if (!cols.some((x) => norm(x) === norm(v))) cols.push(v);
  rows.sort(sortKeyCompare);
  cols.sort(sortKeyCompare);
  const out: Scalar[][] = [['', ...cols, 'Total']];
  for (const r of rows) {
    const line: Scalar[] = [r];
    for (const c of cols) {
      const sel = vals.filter((_, i) => norm(rk[i]) === norm(r) && norm(ck[i]) === norm(c));
      line.push(sel.length ? aggregateLambda(f, sel, ctx) : '');
    }
    line.push(aggregateLambda(f, vals.filter((_, i) => norm(rk[i]) === norm(r)), ctx));
    out.push(line);
  }
  const total: Scalar[] = ['Total'];
  for (const c of cols) total.push(aggregateLambda(f, vals.filter((_, i) => norm(ck[i]) === norm(c)), ctx));
  total.push(aggregateLambda(f, vals, ctx));
  out.push(total);
  return matrix(out);
});

/* ================================================================ misc */

/** Used by the formula bar to show what a reference text points at. */
export function describeRef(ref: RefValue, ctx: EvalContext): string {
  const name = ctx.sheetName(ref.sheet) ?? '';
  const a = addrName(ref.r1, ref.c1);
  const body = ref.r1 === ref.r2 && ref.c1 === ref.c2 ? a : `${a}:${addrName(ref.r2, ref.c2)}`;
  return `${quoteSheet(name)}!${body}`;
}
