import { translateFormula } from '../engine/refshift';
import { compare, isErr, asErr, makeCriteria, type Scalar } from '../engine/values';
import { formatValue, parseInput } from '../format/numfmt';
import { cellKey, keyCol, keyRow, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { cloneCell, type Sheet } from '../model/workbook';
import type { AutoFilter, Cell, ChartSpec, ColumnFilter, CondFormat, DefinedName, ImageSpec, Validation } from '../model/types';
import type { SheetDoc } from '../doc';

/* ======================================================= current region */

/** The block of non-empty cells around a cell (Ctrl+A / auto-detect for sort, filter, charts). */
export function currentRegion(doc: SheetDoc, sheet: Sheet, r: number, c: number): Range {
  const filled = (rr: number, cc: number) => {
    if (rr < 0 || cc < 0) return false;
    const v = doc.value(sheet, rr, cc);
    return v !== undefined && v !== null && v !== '';
  };
  let rg: Range = { r1: r, c1: c, r2: r, c2: c };
  for (let guard = 0; guard < 200; guard++) {
    let grown = false;
    // expand while any neighbour row/column of the block has content
    const rowHas = (rr: number) => {
      for (let cc = Math.max(0, rg.c1 - 1); cc <= rg.c2 + 1; cc++) if (filled(rr, cc)) return true;
      return false;
    };
    const colHas = (cc: number) => {
      for (let rr = Math.max(0, rg.r1 - 1); rr <= rg.r2 + 1; rr++) if (filled(rr, cc)) return true;
      return false;
    };
    if (rg.r1 > 0 && rowHas(rg.r1 - 1)) {
      rg = { ...rg, r1: rg.r1 - 1 };
      grown = true;
    }
    if (rg.r2 < MAX_ROWS - 1 && rowHas(rg.r2 + 1)) {
      rg = { ...rg, r2: rg.r2 + 1 };
      grown = true;
    }
    if (rg.c1 > 0 && colHas(rg.c1 - 1)) {
      rg = { ...rg, c1: rg.c1 - 1 };
      grown = true;
    }
    if (rg.c2 < MAX_COLS - 1 && colHas(rg.c2 + 1)) {
      rg = { ...rg, c2: rg.c2 + 1 };
      grown = true;
    }
    if (!grown) break;
    if ((rg.r2 - rg.r1) * (rg.c2 - rg.c1) > 5_000_000) break;
  }
  return rg;
}

/** Selection to act on: the selected block, or the current region when a single cell is selected. */
export function dataRange(doc: SheetDoc): Range {
  const sel = doc.sel;
  const rg = sel.ranges[sel.ranges.length - 1];
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return currentRegion(doc, doc.sheet, rg.r1, rg.c1);
  return doc.clampToUsed(doc.sheet, rg);
}

/** Guesses whether the first row of a range is a header row. */
export function looksLikeHeader(doc: SheetDoc, sheet: Sheet, rg: Range): boolean {
  if (rg.r2 <= rg.r1) return false;
  let textTop = 0;
  let numBelow = 0;
  let boldTop = 0;
  for (let c = rg.c1; c <= rg.c2; c++) {
    const top = doc.value(sheet, rg.r1, c);
    const below = doc.value(sheet, rg.r1 + 1, c);
    if (typeof top === 'string') textTop++;
    if (typeof below === 'number') numBelow++;
    if (doc.styleOf(sheet, rg.r1, c).bold) boldTop++;
  }
  const cols = rg.c2 - rg.c1 + 1;
  return textTop === cols && (numBelow > 0 || boldTop > 0 || cols === 1);
}

/* ================================================================ sort */

export interface SortKey {
  col: number; // absolute column
  desc?: boolean;
  by?: 'value' | 'fill' | 'color';
  color?: string;
  custom?: string[];
}

export function sortRange(doc: SheetDoc, rg: Range, keys: SortKey[], hasHeader: boolean, opts: { caseSensitive?: boolean; byColumns?: boolean } = {}): void {
  const sheet = doc.sheet;
  const r1 = hasHeader ? rg.r1 + 1 : rg.r1;
  if (r1 > rg.r2 || !keys.length) return;
  const rows: { src: number; cells: (Cell | undefined)[]; keys: Scalar[] }[] = [];
  for (let r = r1; r <= rg.r2; r++) {
    if (sheet.filterHidden.has(r)) continue;
    const cells: (Cell | undefined)[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) cells.push(cloneCell(sheet.get(r, c)));
    const ks = keys.map((k) => {
      if (k.by === 'fill') return (doc.styleOf(sheet, r, k.col).fill ?? '') === (k.color ?? '') ? 0 : 1;
      if (k.by === 'color') return (doc.styleOf(sheet, r, k.col).color ?? '') === (k.color ?? '') ? 0 : 1;
      const v = doc.value(sheet, r, k.col) ?? null;
      if (k.custom && typeof v === 'string') {
        const i = k.custom.findIndex((x) => x.toLowerCase() === v.toLowerCase());
        return i < 0 ? k.custom.length + 1 : i;
      }
      return v;
    });
    rows.push({ src: r, cells, keys: ks });
  }
  const visibleRows = rows.map((x) => x.src);
  rows.sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      const x = a.keys[i];
      const y = b.keys[i];
      // blanks always last, errors before blanks
      const bx = x === null || x === '';
      const by = y === null || y === '';
      if (bx || by) {
        if (bx && by) continue;
        return bx ? 1 : -1;
      }
      let c: number;
      if (isErr(x) || isErr(y)) c = isErr(x) && isErr(y) ? 0 : isErr(x) ? 1 : -1;
      else if (opts.caseSensitive && typeof x === 'string' && typeof y === 'string') c = x < y ? -1 : x > y ? 1 : 0;
      else c = compare(x, y);
      if (c) return keys[i].desc ? -c : c;
    }
    return a.src - b.src;
  });
  doc.transact('Sort', (tx) => {
    rows.forEach((row, i) => {
      const dest = visibleRows[i];
      row.cells.forEach((cell, j) => {
        const c = rg.c1 + j;
        let next = cell;
        if (next?.f !== undefined && dest !== row.src) next = { ...next, f: translateFormula(next.f, dest - row.src, 0), v: undefined };
        tx.set(sheet, cellKey(dest, c), next);
      });
    });
    // row heights travel with their rows
    const heights = rows.map((row) => sheet.rows.get(row.src)?.h);
    if (heights.some((h) => h !== undefined)) {
      tx.prop(sheet, 'rows');
      rows.forEach((_, i) => {
        const dest = visibleRows[i];
        const info = { ...(sheet.rows.get(dest) ?? {}), h: heights[i] };
        if (info.h === undefined) delete info.h;
        sheet.rows.set(dest, info);
      });
    }
  });
}

/* ============================================================== filter */

export function filterValueText(doc: SheetDoc, sheet: Sheet, r: number, c: number): string {
  return doc.displayText(sheet, r, c).text;
}

function passes(doc: SheetDoc, sheet: Sheet, r: number, c: number, f: ColumnFilter, colValues: Scalar[]): boolean {
  const v = doc.value(sheet, r, c) ?? null;
  const text = filterValueText(doc, sheet, r, c);
  if (f.values) {
    const blank = text === '';
    if (blank ? !f.blanks && !f.values.includes('') : !f.values.includes(text)) return false;
  }
  if (f.condition) {
    const { op, value, value2 } = f.condition;
    const low = text.toLowerCase();
    const q = value.toLowerCase();
    switch (op) {
      case 'contains':
        if (!low.includes(q)) return false;
        break;
      case 'notContains':
        if (low.includes(q)) return false;
        break;
      case 'begins':
        if (!low.startsWith(q)) return false;
        break;
      case 'ends':
        if (!low.endsWith(q)) return false;
        break;
      case 'between':
      case 'notBetween': {
        const a = parseInput(value).value;
        const b = parseInput(value2 ?? '').value;
        const inside = typeof v === 'number' && typeof a === 'number' && typeof b === 'number' && v >= Math.min(a, b) && v <= Math.max(a, b);
        if (op === 'between' ? !inside : inside) return false;
        break;
      }
      default: {
        const map: Record<string, string> = { equal: '=', notEqual: '<>', greater: '>', less: '<', greaterEqual: '>=', lessEqual: '<=' };
        if (!makeCriteria(`${map[op]}${value}`)(v)) return false;
      }
    }
  }
  if (f.top) {
    const nums = colValues.filter((x): x is number => typeof x === 'number').sort((a, b) => b - a);
    if (typeof v !== 'number') return false;
    const n = f.top.percent ? Math.max(1, Math.round((nums.length * f.top.n) / 100)) : f.top.n;
    const list = f.top.bottom ? nums.slice(-n) : nums.slice(0, n);
    if (!list.includes(v)) return false;
  }
  if (f.color && (doc.styleOf(sheet, r, c).fill ?? '') !== f.color) return false;
  return true;
}

/** Recomputes which rows the auto-filter hides. */
export function applyFilter(doc: SheetDoc, sheet: Sheet): void {
  sheet.filterHidden.clear();
  const af = sheet.filter;
  if (!af) return;
  const entries = Object.entries(af.columns).filter(([, f]) => f && (f.values || f.condition || f.top || f.color));
  if (!entries.length) return;
  const colValues = new Map<number, Scalar[]>();
  for (const [k] of entries) {
    const c = af.range.c1 + Number(k);
    const vals: Scalar[] = [];
    for (let r = af.range.r1 + 1; r <= af.range.r2; r++) vals.push(doc.value(sheet, r, c) ?? null);
    colValues.set(c, vals);
  }
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) {
    for (const [k, f] of entries) {
      const c = af.range.c1 + Number(k);
      if (!passes(doc, sheet, r, c, f, colValues.get(c)!)) {
        sheet.filterHidden.add(r);
        break;
      }
    }
  }
}

export function toggleFilter(doc: SheetDoc): void {
  const sheet = doc.sheet;
  doc.transact(sheet.filter ? 'Remove filter' : 'Filter', (tx) => {
    tx.prop(sheet, 'filter');
    if (sheet.filter) sheet.filter = undefined;
    else {
      const rg = dataRange(doc);
      sheet.filter = { range: rg, columns: {} };
    }
  });
  applyFilter(doc, sheet);
  doc.emit('edit');
}

export function setColumnFilter(doc: SheetDoc, col: number, f: ColumnFilter | null): void {
  const sheet = doc.sheet;
  const af = sheet.filter;
  if (!af) return;
  doc.transact('Filter', (tx) => {
    tx.prop(sheet, 'filter');
    const columns = { ...af.columns };
    if (f) columns[col - af.range.c1] = f;
    else delete columns[col - af.range.c1];
    // grow the range to include new rows typed below
    const ext = doc.engine.extent(sheet.id);
    let r2 = af.range.r2;
    while (r2 < ext.rows && doc.value(sheet, r2 + 1, af.range.c1) !== undefined) r2++;
    sheet.filter = { ...af, range: { ...af.range, r2 }, columns } as AutoFilter;
  });
  applyFilter(doc, sheet);
  doc.emit('edit');
}

export function clearFilters(doc: SheetDoc): void {
  const sheet = doc.sheet;
  if (!sheet.filter) return;
  doc.transact('Clear filter', (tx) => {
    tx.prop(sheet, 'filter');
    sheet.filter = { ...sheet.filter!, columns: {} };
  });
  applyFilter(doc, sheet);
  doc.emit('edit');
}

/** Distinct display values of a filter column (for the filter menu). */
export function filterChoices(doc: SheetDoc, sheet: Sheet, col: number): { text: string; count: number }[] {
  const af = sheet.filter;
  if (!af) return [];
  const map = new Map<string, number>();
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) {
    const t = filterValueText(doc, sheet, r, col);
    map.set(t, (map.get(t) ?? 0) + 1);
  }
  return [...map.entries()]
    .map(([text, count]) => ({ text, count }))
    .sort((a, b) => {
      if (a.text === '') return 1;
      if (b.text === '') return -1;
      const na = parseFloat(a.text.replace(/[^\d.-]/g, ''));
      const nb = parseFloat(b.text.replace(/[^\d.-]/g, ''));
      if (!Number.isNaN(na) && !Number.isNaN(nb) && /^[\d$€£¥.,%\s-]+$/.test(a.text + b.text)) return na - nb;
      return a.text.localeCompare(b.text, undefined, { numeric: true, sensitivity: 'base' });
    });
}

/* ======================================================== data tools */

export function removeDuplicates(doc: SheetDoc, rg: Range, cols: number[], hasHeader: boolean): { removed: number; kept: number } {
  const sheet = doc.sheet;
  const r1 = hasHeader ? rg.r1 + 1 : rg.r1;
  const seen = new Set<string>();
  const keep: number[] = [];
  for (let r = r1; r <= rg.r2; r++) {
    const key = cols.map((c) => {
      const v = doc.value(sheet, r, c);
      return typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + ':' + String(isErr(v) ? asErr(v).error : v);
    }).join('\u0001');
    if (seen.has(key)) continue;
    seen.add(key);
    keep.push(r);
  }
  const removed = rg.r2 - r1 + 1 - keep.length;
  if (!removed) return { removed: 0, kept: keep.length };
  doc.transact('Remove duplicates', (tx) => {
    const rows = keep.map((r) => {
      const cells: (Cell | undefined)[] = [];
      for (let c = rg.c1; c <= rg.c2; c++) cells.push(cloneCell(sheet.get(r, c)));
      return { r, cells };
    });
    for (let r = r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) tx.set(sheet, cellKey(r, c), undefined);
    rows.forEach((row, i) =>
      row.cells.forEach((cell, j) => {
        let next = cell;
        if (next?.f !== undefined) next = { ...next, f: translateFormula(next.f, r1 + i - row.r, 0), v: undefined };
        tx.set(sheet, cellKey(r1 + i, rg.c1 + j), next);
      }),
    );
  });
  return { removed, kept: keep.length };
}

export function textToColumns(doc: SheetDoc, rg: Range, opts: { delimiters: string[]; other?: string; treatConsecutive?: boolean; quote?: string; fixedWidths?: number[] }): number {
  const sheet = doc.sheet;
  const c0 = rg.c1;
  let maxCols = 0;
  const parsed: { r: number; parts: string[] }[] = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const text = doc.displayText(sheet, r, c0).text;
    if (!text) continue;
    let parts: string[];
    if (opts.fixedWidths?.length) {
      parts = [];
      let pos = 0;
      for (const w of [...opts.fixedWidths, Infinity]) {
        parts.push(text.slice(pos, pos + w).trim());
        pos += w;
        if (pos >= text.length) break;
      }
    } else {
      const delims = [...opts.delimiters, ...(opts.other ? [opts.other] : [])].map((d) => (d === 'tab' ? '\t' : d === 'space' ? ' ' : d));
      const re = new RegExp(`(?:${delims.map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})${opts.treatConsecutive ? '+' : ''}`);
      parts = [];
      const q = opts.quote ?? '"';
      let cur = '';
      let inQ = false;
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (q && ch === q) {
          inQ = !inQ;
          continue;
        }
        if (!inQ) {
          const m = re.exec(text.slice(i));
          if (m && m.index === 0) {
            parts.push(cur);
            cur = '';
            i += m[0].length - 1;
            continue;
          }
        }
        cur += ch;
      }
      parts.push(cur);
    }
    maxCols = Math.max(maxCols, parts.length);
    parsed.push({ r, parts });
  }
  doc.transact('Text to columns', (tx) => {
    for (const { r, parts } of parsed)
      parts.forEach((p, j) => {
        const key = cellKey(r, c0 + j);
        const cur = sheet.cells.get(key);
        const v = parseInput(p).value;
        tx.set(sheet, key, { ...(cur ?? {}), v: v === null ? undefined : v, f: undefined });
      });
  });
  return maxCols;
}

/* ========================================================= find/replace */

export interface FindOptions {
  matchCase?: boolean;
  wholeCell?: boolean;
  regex?: boolean;
  lookIn?: 'values' | 'formulas';
  scope?: 'sheet' | 'workbook' | 'selection';
}

export interface FindHit {
  sheet: number;
  r: number;
  c: number;
  text: string;
}

function matcher(query: string, o: FindOptions): RegExp | null {
  try {
    let src = o.regex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\\\*/g, '.*').replace(/\\\?/g, '.');
    if (o.wholeCell) src = `^(?:${src})$`;
    return new RegExp(src, o.matchCase ? 'g' : 'gi');
  } catch {
    return null;
  }
}

export function findAll(doc: SheetDoc, query: string, o: FindOptions = {}): FindHit[] {
  if (!query) return [];
  const re = matcher(query, o);
  if (!re) return [];
  const sheets = o.scope === 'workbook' ? doc.wb.sheets : [doc.sheet];
  const hits: FindHit[] = [];
  const ranges = o.scope === 'selection' ? doc.sel.ranges : null;
  for (const sheet of sheets) {
    const keys = [...sheet.cells.keys()].sort((a, b) => keyRow(a) - keyRow(b) || keyCol(a) - keyCol(b));
    for (const k of keys) {
      const r = keyRow(k);
      const c = keyCol(k);
      if (ranges && !ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2)) continue;
      const cell = sheet.cells.get(k)!;
      const text = o.lookIn === 'formulas' && cell.f !== undefined ? '=' + cell.f : doc.displayText(sheet, r, c).text;
      if (!text) continue;
      re.lastIndex = 0;
      if (re.test(text)) hits.push({ sheet: sheet.id, r, c, text });
      if (hits.length > 20000) return hits;
    }
  }
  return hits;
}

/** Replaces in formulas/constants (not in computed values). Returns the number of cells changed. */
export function replaceAll(doc: SheetDoc, query: string, replacement: string, o: FindOptions = {}, only?: FindHit[]): number {
  const re = matcher(query, o);
  if (!re) return 0;
  const hits = only ?? findAll(doc, query, { ...o, lookIn: 'formulas' });
  let n = 0;
  doc.transact('Replace', (tx) => {
    for (const h of hits) {
      const sheet = doc.wb.sheetById(h.sheet);
      if (!sheet) continue;
      const cell = sheet.get(h.r, h.c);
      if (!cell) continue;
      if (cell.f !== undefined) {
        re.lastIndex = 0;
        const nf = cell.f.replace(re, replacement);
        if (nf !== cell.f) {
          try {
            const parsed = doc.interpret('=' + nf, sheet);
            tx.patch(sheet, cellKey(h.r, h.c), { f: parsed.cell.f, v: undefined });
            n++;
          } catch {
            /* skip replacements that would break the formula */
          }
        }
        continue;
      }
      const t = typeof cell.v === 'string' ? cell.v : cell.v === undefined || cell.v === null ? '' : formatValue(cell.v, doc.wb.style(cell.s).numFmt).text;
      re.lastIndex = 0;
      const nt = t.replace(re, replacement);
      if (nt !== t) {
        const v = typeof cell.v === 'string' ? nt : parseInput(nt).value;
        tx.patch(sheet, cellKey(h.r, h.c), { v: v === null ? undefined : v });
        n++;
      }
    }
  });
  return n;
}

/* ================================================== model collections */

let seq = 0;
export const newId = (p: string) => `${p}${Date.now().toString(36)}${(seq++).toString(36)}`;

export function upsertCF(doc: SheetDoc, rule: CondFormat): void {
  const sheet = doc.sheet;
  doc.transact('Conditional formatting', (tx) => {
    tx.prop(sheet, 'cf');
    const i = sheet.cf.findIndex((x) => x.id === rule.id);
    if (i >= 0) sheet.cf[i] = rule;
    else sheet.cf.push(rule);
  });
}

export function deleteCF(doc: SheetDoc, ids: string[] | 'selection' | 'sheet'): void {
  const sheet = doc.sheet;
  doc.transact('Clear rules', (tx) => {
    tx.prop(sheet, 'cf');
    if (ids === 'sheet') sheet.cf = [];
    else if (ids === 'selection') {
      const sel = doc.sel.ranges;
      sheet.cf = sheet.cf.filter((cf) => !cf.ranges.some((r) => sel.some((s) => r.r1 <= s.r2 && r.r2 >= s.r1 && r.c1 <= s.c2 && r.c2 >= s.c1)));
    } else sheet.cf = sheet.cf.filter((cf) => !ids.includes(cf.id));
  });
}

export function reorderCF(doc: SheetDoc, id: string, delta: number): void {
  const sheet = doc.sheet;
  const i = sheet.cf.findIndex((x) => x.id === id);
  const j = i + delta;
  if (i < 0 || j < 0 || j >= sheet.cf.length) return;
  doc.transact('Rule priority', (tx) => {
    tx.prop(sheet, 'cf');
    const [r] = sheet.cf.splice(i, 1);
    sheet.cf.splice(j, 0, r);
  });
}

export function setValidation(doc: SheetDoc, v: Validation | null, ranges = doc.sel.ranges): void {
  const sheet = doc.sheet;
  doc.transact(v ? 'Data validation' : 'Clear validation', (tx) => {
    tx.prop(sheet, 'validations');
    // remove existing rules from the target ranges (splitting is approximated by dropping overlaps)
    sheet.validations = sheet.validations.filter((x) => !x.ranges.some((r) => ranges.some((s) => r.r1 >= s.r1 && r.r2 <= s.r2 && r.c1 >= s.c1 && r.c2 <= s.c2)));
    for (const x of sheet.validations) x.ranges = x.ranges.filter((r) => !ranges.some((s) => r.r1 <= s.r2 && r.r2 >= s.r1 && r.c1 <= s.c2 && r.c2 >= s.c1));
    sheet.validations = sheet.validations.filter((x) => x.ranges.length);
    if (v) sheet.validations.push({ ...v, ranges: ranges.map((r) => doc.clampToUsed(sheet, r, 1000)) });
  });
}

export function validationAt(sheet: Sheet, r: number, c: number): Validation | undefined {
  return sheet.validations.find((v) => v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
}

export function setNames(doc: SheetDoc, names: DefinedName[]): void {
  doc.transact('Names', (tx) => {
    tx.names();
    doc.wb.names = names;
  });
}

export function upsertChart(doc: SheetDoc, chart: ChartSpec, sheet = doc.sheet): void {
  doc.transact('Chart', (tx) => {
    tx.prop(sheet, 'charts');
    const i = sheet.charts.findIndex((x) => x.id === chart.id);
    if (i >= 0) sheet.charts[i] = chart;
    else sheet.charts.push(chart);
  });
}

export function deleteChart(doc: SheetDoc, id: string, sheet = doc.sheet): void {
  doc.transact('Delete chart', (tx) => {
    tx.prop(sheet, 'charts');
    sheet.charts = sheet.charts.filter((c) => c.id !== id);
  });
}

export function upsertImage(doc: SheetDoc, img: ImageSpec, sheet = doc.sheet): void {
  doc.transact('Picture', (tx) => {
    tx.prop(sheet, 'images');
    const i = sheet.images.findIndex((x) => x.id === img.id);
    if (i >= 0) sheet.images[i] = img;
    else sheet.images.push(img);
  });
}

export function deleteImage(doc: SheetDoc, id: string, sheet = doc.sheet): void {
  doc.transact('Delete picture', (tx) => {
    tx.prop(sheet, 'images');
    sheet.images = sheet.images.filter((c) => c.id !== id);
  });
}

export function setCellMeta(doc: SheetDoc, r: number, c: number, patch: { note?: string | undefined; link?: string | undefined }): void {
  const sheet = doc.sheet;
  doc.transact(patch.note !== undefined || 'note' in patch ? 'Note' : 'Link', (tx) => tx.patch(sheet, cellKey(r, c), patch));
}
