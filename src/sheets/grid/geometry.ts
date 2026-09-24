import { MAX_COLS, MAX_ROWS } from '../model/address';
import type { Sheet } from '../model/workbook';

/**
 * Positions along one axis with sparse custom sizes.
 * position(i) = i*def + Σ(delta of custom entries before i); O(log n) both ways.
 */
export class Axis {
  private idx: number[] = [];
  private size: number[] = [];
  private cum: number[] = []; // cumulative delta *before* entry k
  constructor(
    public readonly def: number,
    public readonly count: number,
    sizes: Map<number, number>,
  ) {
    const entries = [...sizes.entries()].filter(([i, s]) => i < count && s !== def).sort((a, b) => a[0] - b[0]);
    let acc = 0;
    for (const [i, s] of entries) {
      this.idx.push(i);
      this.size.push(s);
      this.cum.push(acc);
      acc += s - def;
    }
  }

  sizeOf(i: number): number {
    const k = this.find(i);
    return k >= 0 && this.idx[k] === i ? this.size[k] : this.def;
  }

  /** Index of the last custom entry with idx <= i (or -1). */
  private find(i: number): number {
    let lo = 0;
    let hi = this.idx.length - 1;
    let best = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (this.idx[mid] <= i) {
        best = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    return best;
  }

  /** Start position of index i. */
  pos(i: number): number {
    const k = this.find(i - 1);
    if (k < 0) return i * this.def;
    return i * this.def + this.cum[k] + (this.size[k] - this.def);
  }

  total(): number {
    return this.pos(this.count);
  }

  /** Index containing position p (skips zero-size entries). */
  at(p: number): number {
    if (p <= 0) return this.firstVisible(0);
    // binary search over custom entries by their start position
    let lo = 0;
    let hi = this.idx.length - 1;
    let k = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const start = this.idx[mid] * this.def + this.cum[mid];
      if (start <= p) {
        k = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    let i: number;
    if (k < 0) i = Math.floor(p / this.def);
    else {
      const start = this.idx[k] * this.def + this.cum[k];
      const end = start + this.size[k];
      if (p < end) i = this.idx[k];
      else i = this.idx[k] + 1 + Math.floor((p - end) / this.def);
      // skip over later custom entries that the linear estimate jumped past
      while (k + 1 < this.idx.length && this.idx[k + 1] <= i) {
        k++;
        const s2 = this.idx[k] * this.def + this.cum[k];
        const e2 = s2 + this.size[k];
        if (p < s2) {
          i = this.idx[k] - 1;
          break;
        }
        if (p < e2) {
          i = this.idx[k];
          break;
        }
        i = this.idx[k] + 1 + Math.floor((p - e2) / this.def);
      }
    }
    return Math.min(this.count - 1, Math.max(0, i));
  }

  firstVisible(from: number): number {
    let i = from;
    while (i < this.count - 1 && this.sizeOf(i) === 0) i++;
    return i;
  }

  /** Next visible index in a direction (skipping hidden rows/cols). */
  step(i: number, dir: 1 | -1): number {
    let j = i + dir;
    while (j >= 0 && j < this.count && this.sizeOf(j) === 0) j += dir;
    if (j < 0 || j >= this.count) return i;
    return j;
  }
}

export interface SheetGeometry {
  rows: Axis;
  cols: Axis;
}

const cache = new WeakMap<Sheet, { key: string; geo: SheetGeometry }>();

/** Row/column axes for a sheet (cached until sizes change). */
export function geometryOf(sheet: Sheet, version: number): SheetGeometry {
  const key = `${version}:${sheet.defaultRowHeight}:${sheet.defaultColWidth}:${sheet.filterHidden.size}`;
  const hit = cache.get(sheet);
  if (hit && hit.key === key) return hit.geo;
  const rowSizes = new Map<number, number>();
  for (const [r, info] of sheet.rows) {
    const h = info.hidden ? 0 : info.h;
    if (h !== undefined) rowSizes.set(r, h);
  }
  for (const r of sheet.filterHidden) rowSizes.set(r, 0);
  const colSizes = new Map<number, number>();
  for (const [c, info] of sheet.cols) {
    const w = info.hidden ? 0 : info.w;
    if (w !== undefined) colSizes.set(c, w);
  }
  const geo = { rows: new Axis(sheet.defaultRowHeight, MAX_ROWS, rowSizes), cols: new Axis(sheet.defaultColWidth, MAX_COLS, colSizes) };
  cache.set(sheet, { key, geo });
  return geo;
}
