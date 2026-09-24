import type { Scalar } from '../engine/values';
import { FErr, isErr, asErr } from '../engine/values';
import { cellKey, keyCol, keyRow, MAX_COLS, type Range } from './address';
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  type AutoFilter,
  type Cell,
  type CellStyle,
  type ChartSpec,
  type ColInfo,
  type CondFormat,
  type DefinedName,
  type ImageSpec,
  type RowInfo,
  type SheetPrint,
  type SheetProtection,
  type SheetView,
  type Validation,
} from './types';

export class Sheet {
  id: number;
  name: string;
  cells = new Map<number, Cell>();
  cols = new Map<number, ColInfo>();
  rows = new Map<number, RowInfo>();
  merges: Range[] = [];
  cf: CondFormat[] = [];
  validations: Validation[] = [];
  charts: ChartSpec[] = [];
  images: ImageSpec[] = [];
  filter?: AutoFilter;
  view: SheetView = { zoom: 1, showGrid: true, showHeaders: true, showFormulas: false, rtl: false, freezeRows: 0, freezeCols: 0, scrollRow: 0, scrollCol: 0 };
  protection?: SheetProtection;
  print?: SheetPrint;
  tabColor?: string;
  hidden = false;
  /** Rows currently hidden by the auto-filter (recomputed when the filter is applied). */
  filterHidden = new Set<number>();
  defaultColWidth = DEFAULT_COL_WIDTH;
  defaultRowHeight = DEFAULT_ROW_HEIGHT;

  private maxR = -1;
  private maxC = -1;
  private extentStale = false;

  constructor(id: number, name: string) {
    this.id = id;
    this.name = name;
  }

  get(r: number, c: number): Cell | undefined {
    return this.cells.get(r * MAX_COLS + c);
  }

  /** Replaces a cell (undefined removes it). */
  put(key: number, cell: Cell | undefined): void {
    if (!cell || isEmptyCell(cell)) {
      if (this.cells.delete(key)) this.extentStale = true;
      return;
    }
    this.cells.set(key, cell);
    const r = keyRow(key);
    const c = keyCol(key);
    if (!this.extentStale) {
      if (r > this.maxR) this.maxR = r;
      if (c > this.maxC) this.maxC = c;
    }
  }

  /** Last used row/column indexes (-1 when empty). */
  extent(): { rows: number; cols: number } {
    if (this.extentStale) {
      let mr = -1;
      let mc = -1;
      for (const k of this.cells.keys()) {
        const r = keyRow(k);
        const c = keyCol(k);
        if (r > mr) mr = r;
        if (c > mc) mc = c;
      }
      this.maxR = mr;
      this.maxC = mc;
      this.extentStale = false;
    }
    return { rows: this.maxR, cols: this.maxC };
  }

  markExtentStale(): void {
    this.extentStale = true;
  }

  colWidth(c: number): number {
    const ci = this.cols.get(c);
    if (ci?.hidden) return 0;
    return ci?.w ?? this.defaultColWidth;
  }

  rowHeight(r: number): number {
    const ri = this.rows.get(r);
    if (ri?.hidden || this.filterHidden.has(r)) return 0;
    return ri?.h ?? this.defaultRowHeight;
  }

  isRowHidden(r: number): boolean {
    return !!this.rows.get(r)?.hidden || this.filterHidden.has(r);
  }

  isColHidden(c: number): boolean {
    return !!this.cols.get(c)?.hidden;
  }

  /** The merge containing a cell, if any. */
  mergeAt(r: number, c: number): Range | undefined {
    for (const m of this.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return undefined;
  }
}

export function isEmptyCell(c: Cell): boolean {
  return (c.v === undefined || c.v === null) && c.f === undefined && !c.s && !c.link && !c.note;
}

export function cloneCell(c: Cell | undefined): Cell | undefined {
  if (!c) return undefined;
  const out: Cell = {};
  if (c.v !== undefined) out.v = c.v;
  if (c.f !== undefined) out.f = c.f;
  if (c.s) out.s = c.s;
  if (c.link) out.link = c.link;
  if (c.note) out.note = c.note;
  return out;
}

/** Canonical JSON for style interning (sorted keys, no undefined). */
function styleKey(s: CellStyle): string {
  const keys = Object.keys(s).sort() as (keyof CellStyle)[];
  const o: Record<string, unknown> = {};
  for (const k of keys) {
    const v = s[k];
    if (v === undefined || v === null || v === '' || (v === false && k !== 'locked')) continue;
    if (k === 'locked' && v === true) continue; // cells are locked by default
    o[k] = v;
  }
  return JSON.stringify(o);
}

export function cleanStyle(s: CellStyle): CellStyle {
  return JSON.parse(styleKey(s));
}

export class Workbook {
  sheets: Sheet[] = [];
  styles: CellStyle[] = [{}];
  names: DefinedName[] = [];
  props: { title?: string; author?: string; created?: number; company?: string } = {};
  calcAuto = true;
  activeSheet = 0;
  private styleIndex = new Map<string, number>([['{}', 0]]);
  private nextId = 1;

  addSheet(name?: string, index?: number): Sheet {
    const sheet = new Sheet(this.nextId++, name ?? this.uniqueSheetName('Sheet'));
    if (index === undefined) this.sheets.push(sheet);
    else this.sheets.splice(index, 0, sheet);
    return sheet;
  }

  /** Re-inserts an existing sheet object (undo of delete). */
  insertSheet(sheet: Sheet, index: number): void {
    this.sheets.splice(index, 0, sheet);
    if (sheet.id >= this.nextId) this.nextId = sheet.id + 1;
  }

  sheetById(id: number): Sheet | undefined {
    return this.sheets.find((s) => s.id === id);
  }

  sheetByName(name: string): Sheet | undefined {
    const low = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === low);
  }

  uniqueSheetName(base: string): string {
    const taken = new Set(this.sheets.map((s) => s.name.toLowerCase()));
    if (base !== 'Sheet' && !taken.has(base.toLowerCase())) return base;
    for (let i = 1; ; i++) {
      const n = base === 'Sheet' ? `Sheet${i}` : `${base} (${i + 1})`;
      if (!taken.has(n.toLowerCase())) return n;
    }
  }

  styleId(style: CellStyle): number {
    const key = styleKey(style);
    let id = this.styleIndex.get(key);
    if (id === undefined) {
      id = this.styles.length;
      this.styles.push(JSON.parse(key));
      this.styleIndex.set(key, id);
    }
    return id;
  }

  style(id: number | undefined): CellStyle {
    return (id && this.styles[id]) || this.styles[0];
  }

  /** Style id after merging a patch (undefined values remove properties). */
  patchStyle(id: number | undefined, patch: Partial<CellStyle>): number {
    const base = { ...this.style(id) } as Record<string, unknown>;
    for (const [k, v] of Object.entries(patch)) {
      if (v === undefined || v === null) delete base[k];
      else base[k] = v;
    }
    return this.styleId(base as CellStyle);
  }

  /** Effective style of a cell: cell style, else row style, else column style. */
  cellStyleId(sheet: Sheet, r: number, c: number): number {
    const cell = sheet.cells.get(cellKey(r, c));
    if (cell?.s !== undefined) return cell.s;
    const rs = sheet.rows.get(r)?.s;
    if (rs) return rs;
    return sheet.cols.get(c)?.s ?? 0;
  }
}

/* ============================================================ serialise */

export interface WorkbookJSON {
  format: 'affice-sheet';
  version: 1;
  props: Workbook['props'];
  styles: CellStyle[];
  names: DefinedName[];
  active: number;
  calcAuto: boolean;
  sheets: SheetJSON[];
}

export interface SheetJSON {
  id: number;
  name: string;
  cells: [number, number, Cell][];
  cols: [number, ColInfo][];
  rows: [number, RowInfo][];
  merges: Range[];
  cf: CondFormat[];
  validations: Validation[];
  charts: ChartSpec[];
  images: ImageSpec[];
  filter?: AutoFilter;
  view: SheetView;
  protection?: SheetProtection;
  print?: SheetPrint;
  tabColor?: string;
  hidden?: boolean;
  defaultColWidth?: number;
  defaultRowHeight?: number;
}

function encodeValue(v: Scalar | undefined): unknown {
  if (isErr(v)) return { error: asErr(v).error };
  return v;
}

function decodeValue(v: unknown): Scalar | undefined {
  if (v && typeof v === 'object' && 'error' in v) return new FErr((v as { error: FErr['error'] }).error);
  return v as Scalar | undefined;
}

export function sheetToJSON(s: Sheet): SheetJSON {
  const cells: [number, number, Cell][] = [];
  for (const [k, cell] of s.cells) {
    const c = cloneCell(cell)!;
    if (c.v !== undefined) c.v = encodeValue(c.v) as Scalar;
    cells.push([keyRow(k), keyCol(k), c]);
  }
  return {
    id: s.id,
    name: s.name,
    cells,
    cols: [...s.cols.entries()],
    rows: [...s.rows.entries()],
    merges: s.merges,
    cf: s.cf,
    validations: s.validations,
    charts: s.charts,
    images: s.images,
    filter: s.filter,
    view: s.view,
    protection: s.protection,
    print: s.print,
    tabColor: s.tabColor,
    hidden: s.hidden || undefined,
    defaultColWidth: s.defaultColWidth,
    defaultRowHeight: s.defaultRowHeight,
  };
}

export function sheetFromJSON(j: SheetJSON): Sheet {
  const s = new Sheet(j.id, j.name);
  for (const [r, c, cell] of j.cells) {
    const copy = { ...cell };
    if (copy.v !== undefined) copy.v = decodeValue(copy.v);
    s.put(cellKey(r, c), copy);
  }
  s.cols = new Map(j.cols);
  s.rows = new Map(j.rows);
  s.merges = j.merges ?? [];
  s.cf = j.cf ?? [];
  s.validations = j.validations ?? [];
  s.charts = j.charts ?? [];
  s.images = j.images ?? [];
  s.filter = j.filter;
  s.view = { ...s.view, ...(j.view ?? {}) };
  s.protection = j.protection;
  s.print = j.print;
  s.tabColor = j.tabColor;
  s.hidden = !!j.hidden;
  if (j.defaultColWidth) s.defaultColWidth = j.defaultColWidth;
  if (j.defaultRowHeight) s.defaultRowHeight = j.defaultRowHeight;
  return s;
}

export function workbookToJSON(wb: Workbook): WorkbookJSON {
  return {
    format: 'affice-sheet',
    version: 1,
    props: wb.props,
    styles: wb.styles,
    names: wb.names,
    active: wb.activeSheet,
    calcAuto: wb.calcAuto,
    sheets: wb.sheets.map(sheetToJSON),
  };
}

export function workbookFromJSON(j: WorkbookJSON): Workbook {
  const wb = new Workbook();
  wb.props = j.props ?? {};
  // rebuild the style table through the interning map so ids stay stable
  wb.styles = [{}];
  const idMap = new Map<number, number>();
  (j.styles ?? [{}]).forEach((st, i) => idMap.set(i, wb.styleId(st)));
  wb.names = j.names ?? [];
  wb.calcAuto = j.calcAuto ?? true;
  for (const sj of j.sheets) {
    const s = sheetFromJSON(sj);
    for (const cell of s.cells.values()) if (cell.s) cell.s = idMap.get(cell.s) ?? 0;
    for (const ci of s.cols.values()) if (ci.s) ci.s = idMap.get(ci.s) ?? 0;
    for (const ri of s.rows.values()) if (ri.s) ri.s = idMap.get(ri.s) ?? 0;
    wb.insertSheet(s, wb.sheets.length);
  }
  wb.activeSheet = Math.min(Math.max(0, j.active ?? 0), wb.sheets.length - 1);
  return wb;
}
