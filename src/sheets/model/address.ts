/** A1-style addressing helpers. Rows and columns are 0-based internally. */

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export interface CellAddr {
  r: number;
  c: number;
}

export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export const cellKey = (r: number, c: number) => r * MAX_COLS + c;
export const keyRow = (k: number) => Math.floor(k / MAX_COLS);
export const keyCol = (k: number) => k % MAX_COLS;

export function colName(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function addrName(r: number, c: number, absR = false, absC = false): string {
  return `${absC ? '$' : ''}${colName(c)}${absR ? '$' : ''}${r + 1}`;
}

export function rangeName(rg: Range): string {
  const a = addrName(rg.r1, rg.c1);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return a;
  if (rg.r1 === 0 && rg.r2 === MAX_ROWS - 1) return `${colName(rg.c1)}:${colName(rg.c2)}`;
  if (rg.c1 === 0 && rg.c2 === MAX_COLS - 1) return `${rg.r1 + 1}:${rg.r2 + 1}`;
  return `${a}:${addrName(rg.r2, rg.c2)}`;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

export function parseAddr(s: string): CellAddr | null {
  const m = CELL_RE.exec(s.trim());
  if (!m) return null;
  const c = colIndex(m[1]);
  const r = Number(m[2]) - 1;
  if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return null;
  return { r, c };
}

/** Parses "A1", "A1:B5", "A:C", "3:7" (without a sheet prefix). */
export function parseRange(s: string): Range | null {
  const t = s.trim().replace(/\$/g, '');
  const parts = t.split(':');
  if (parts.length === 1) {
    const a = parseAddr(parts[0]);
    return a ? { r1: a.r, c1: a.c, r2: a.r, c2: a.c } : null;
  }
  if (parts.length !== 2) return null;
  const [p, q] = parts;
  if (/^[A-Za-z]{1,3}$/.test(p) && /^[A-Za-z]{1,3}$/.test(q)) {
    const c1 = colIndex(p);
    const c2 = colIndex(q);
    return { r1: 0, c1: Math.min(c1, c2), r2: MAX_ROWS - 1, c2: Math.max(c1, c2) };
  }
  if (/^\d+$/.test(p) && /^\d+$/.test(q)) {
    const r1 = Number(p) - 1;
    const r2 = Number(q) - 1;
    return { r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAX_COLS - 1 };
  }
  const a = parseAddr(p);
  const b = parseAddr(q);
  if (!a || !b) return null;
  return normRange({ r1: a.r, c1: a.c, r2: b.r, c2: b.c });
}

export function normRange(r: Range): Range {
  return { r1: Math.min(r.r1, r.r2), c1: Math.min(r.c1, r.c2), r2: Math.max(r.r1, r.r2), c2: Math.max(r.c1, r.c2) };
}

export const inRange = (rg: Range, r: number, c: number) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;

export const rangesIntersect = (a: Range, b: Range) => a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;

export function rangeSize(r: Range): number {
  return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
}

export function sameRange(a: Range, b: Range): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

/** Quotes a sheet name for use in formulas when needed. */
export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}
