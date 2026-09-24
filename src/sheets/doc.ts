import { CalcEngine } from './engine/calc';
import { FormulaError, parseFormula, tokenize } from './engine/parser';
import { isErr, asErr, type Scalar } from './engine/values';
import { translateFormula } from './engine/refshift';
import { formatValue, isDateFormat, parseInput, SHORT_DATE } from './format/numfmt';
import { cellKey, keyCol, keyRow, MAX_COLS, MAX_ROWS, type Range } from './model/address';
import { cloneCell, Sheet, Workbook } from './model/workbook';
import type { Border, BorderStyle, Cell, CellStyle, ColInfo, DefinedName, RowInfo } from './model/types';

export interface CellPos {
  r: number;
  c: number;
}

export interface SheetSelection {
  ranges: Range[];
  active: CellPos;
  anchor: CellPos;
}

export type SheetProp =
  | 'cells'
  | 'cols'
  | 'rows'
  | 'merges'
  | 'cf'
  | 'validations'
  | 'charts'
  | 'images'
  | 'filter'
  | 'protection'
  | 'tabColor'
  | 'hidden'
  | 'name'
  | 'defaultColWidth'
  | 'defaultRowHeight'
  | 'freeze'
  | 'print';

interface CellChange {
  sheet: number;
  key: number;
  before?: Cell;
  after?: Cell;
}

interface PropChange {
  sheet: number;
  prop: SheetProp;
  before: unknown;
  after?: unknown;
}

export interface Transaction {
  label: string;
  cells: Map<string, CellChange>;
  props: Map<string, PropChange>;
  sheetsBefore?: Sheet[];
  sheetsAfter?: Sheet[];
  namesBefore?: DefinedName[];
  namesAfter?: DefinedName[];
  activeBefore: number;
  activeAfter?: number;
  selBefore: { sheet: number; sel: SheetSelection };
  selAfter?: { sheet: number; sel: SheetSelection };
  rebuild: boolean;
}

/* ------------------------------------------------------------ snapshots */

function snapProp(sheet: Sheet, prop: SheetProp): unknown {
  switch (prop) {
    case 'cells':
      return [...sheet.cells.entries()].map(([k, c]) => [k, cloneCell(c)!] as const);
    case 'cols':
      return [...sheet.cols.entries()].map(([k, v]) => [k, { ...v }] as const);
    case 'rows':
      return [...sheet.rows.entries()].map(([k, v]) => [k, { ...v }] as const);
    case 'freeze':
      return { rows: sheet.view.freezeRows, cols: sheet.view.freezeCols };
    default:
      return structuredClone((sheet as unknown as Record<string, unknown>)[prop]);
  }
}

function restoreProp(sheet: Sheet, prop: SheetProp, value: unknown): void {
  switch (prop) {
    case 'cells':
      sheet.cells = new Map((value as [number, Cell][]).map(([k, c]) => [k, cloneCell(c)!]));
      sheet.markExtentStale();
      return;
    case 'cols':
      sheet.cols = new Map((value as [number, ColInfo][]).map(([k, v]) => [k, { ...v }]));
      return;
    case 'rows':
      sheet.rows = new Map((value as [number, RowInfo][]).map(([k, v]) => [k, { ...v }]));
      return;
    case 'freeze': {
      const f = value as { rows: number; cols: number };
      sheet.view.freezeRows = f.rows;
      sheet.view.freezeCols = f.cols;
      return;
    }
    default:
      (sheet as unknown as Record<string, unknown>)[prop] = structuredClone(value);
  }
}

/* ------------------------------------------------------------- builder */

export class TxBuilder {
  readonly data: Transaction;
  constructor(
    private doc: SheetDoc,
    label: string,
  ) {
    this.data = {
      label,
      cells: new Map(),
      props: new Map(),
      activeBefore: doc.wb.activeSheet,
      selBefore: { sheet: doc.sheet.id, sel: structuredClone(doc.sel) },
      rebuild: false,
    };
  }

  get empty(): boolean {
    const d = this.data;
    return !d.cells.size && !d.props.size && !d.sheetsBefore && !d.namesBefore;
  }

  set rebuild(v: boolean) {
    this.data.rebuild = v;
  }

  /** Replaces a cell (undefined removes it). */
  set(sheet: Sheet, key: number, cell: Cell | undefined): void {
    const id = `${sheet.id}:${key}`;
    if (!this.data.cells.has(id)) this.data.cells.set(id, { sheet: sheet.id, key, before: cloneCell(sheet.cells.get(key)) });
    sheet.put(key, cell);
  }

  /** Merges properties into a cell (undefined values delete them). */
  patch(sheet: Sheet, key: number, patch: Partial<Cell>): void {
    const cur = sheet.cells.get(key);
    const next: Cell = { ...(cur ?? {}) };
    for (const [k, v] of Object.entries(patch) as [keyof Cell, unknown][]) {
      if (v === undefined) delete next[k];
      else (next as Record<string, unknown>)[k] = v;
    }
    this.set(sheet, key, next);
  }

  prop(sheet: Sheet, prop: SheetProp): void {
    const id = `${sheet.id}:${prop}`;
    if (this.data.props.has(id)) return;
    this.data.props.set(id, { sheet: sheet.id, prop, before: snapProp(sheet, prop) });
    if (prop === 'cells') this.data.rebuild = true;
  }

  sheets(): void {
    if (!this.data.sheetsBefore) this.data.sheetsBefore = [...this.doc.wb.sheets];
    this.data.rebuild = true;
  }

  names(): void {
    if (!this.data.namesBefore) this.data.namesBefore = structuredClone(this.doc.wb.names);
    this.data.rebuild = true;
  }

  finalize(): void {
    const d = this.data;
    const wb = this.doc.wb;
    for (const ch of d.cells.values()) {
      const s = wb.sheetById(ch.sheet);
      ch.after = cloneCell(s?.cells.get(ch.key));
    }
    for (const ch of d.props.values()) {
      const s = wb.sheetById(ch.sheet);
      if (s) ch.after = snapProp(s, ch.prop);
    }
    if (d.sheetsBefore) d.sheetsAfter = [...wb.sheets];
    if (d.namesBefore) d.namesAfter = structuredClone(wb.names);
    d.activeAfter = wb.activeSheet;
    d.selAfter = { sheet: this.doc.sheet.id, sel: structuredClone(this.doc.sel) };
  }
}

/* ============================================================== the doc */

export interface ClipData {
  sheet: number;
  range: Range;
  cells: (Cell | undefined)[][];
  cut: boolean;
  text: string;
  colWidths: number[];
  merges: Range[];
}

const DATE_FNS: Record<string, string> = {
  TODAY: SHORT_DATE,
  DATE: SHORT_DATE,
  EDATE: SHORT_DATE,
  EOMONTH: SHORT_DATE,
  WORKDAY: SHORT_DATE,
  'WORKDAY.INTL': SHORT_DATE,
  DATEVALUE: SHORT_DATE,
  NOW: `${SHORT_DATE} h:mm`,
  TIME: 'h:mm AM/PM',
  TIMEVALUE: 'h:mm AM/PM',
};

export class SheetDoc {
  wb: Workbook;
  engine: CalcEngine;
  /** Selection per sheet id. */
  private selections = new Map<number, SheetSelection>();
  undoStack: Transaction[] = [];
  redoStack: Transaction[] = [];
  clip: ClipData | null = null;
  version = 0;
  onChange?: (kind: 'edit' | 'view') => void;
  private listeners = new Set<() => void>();

  constructor(wb: Workbook) {
    this.wb = wb;
    if (!wb.sheets.length) wb.addSheet();
    this.engine = new CalcEngine(wb);
    this.engine.rebuild();
  }

  /* ---------------------------------------------------------- events */

  subscribe = (fn: () => void): (() => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  getVersion = () => this.version;

  emit(kind: 'edit' | 'view' = 'view'): void {
    this.version++;
    for (const l of this.listeners) l();
    this.onChange?.(kind);
  }

  /* -------------------------------------------------------- sheets */

  get sheet(): Sheet {
    return this.wb.sheets[this.wb.activeSheet] ?? this.wb.sheets[0];
  }

  get sel(): SheetSelection {
    const id = this.sheet.id;
    let s = this.selections.get(id);
    if (!s) {
      const saved = this.sheet.view.selection;
      s = saved
        ? { ranges: saved.ranges, active: saved.active, anchor: { ...saved.active } }
        : { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 } };
      this.selections.set(id, s);
    }
    return s;
  }

  set sel(v: SheetSelection) {
    this.selections.set(this.sheet.id, v);
    this.sheet.view.selection = { ranges: v.ranges, active: v.active };
  }

  setActiveSheet(index: number): void {
    if (index < 0 || index >= this.wb.sheets.length || index === this.wb.activeSheet) return;
    this.wb.activeSheet = index;
    this.emit('view');
  }

  /* ----------------------------------------------------- transactions */

  transact(label: string, fn: (tx: TxBuilder) => void): boolean {
    const tx = new TxBuilder(this, label);
    fn(tx);
    if (tx.empty) return false;
    tx.finalize();
    this.undoStack.push(tx.data);
    if (this.undoStack.length > 300) this.undoStack.shift();
    this.redoStack = [];
    this.syncEngine(tx.data);
    this.emit('edit');
    return true;
  }

  private syncEngine(tx: Transaction): void {
    if (tx.rebuild || tx.sheetsBefore || tx.namesBefore) {
      this.engine.rebuild();
      return;
    }
    const changes = [...tx.cells.values()].map((c) => ({ sheet: c.sheet, key: c.key }));
    if (changes.length) this.engine.cellsChanged(changes);
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undo(): string | null {
    const tx = this.undoStack.pop();
    if (!tx) return null;
    this.applySide(tx, 'before');
    this.redoStack.push(tx);
    return tx.label;
  }

  redo(): string | null {
    const tx = this.redoStack.pop();
    if (!tx) return null;
    this.applySide(tx, 'after');
    this.undoStack.push(tx);
    return tx.label;
  }

  private applySide(tx: Transaction, side: 'before' | 'after'): void {
    const wb = this.wb;
    if (tx.sheetsBefore) wb.sheets = [...(side === 'before' ? tx.sheetsBefore : tx.sheetsAfter!)];
    if (tx.namesBefore) wb.names = structuredClone(side === 'before' ? tx.namesBefore : tx.namesAfter!);
    for (const ch of tx.props.values()) {
      const s = wb.sheetById(ch.sheet);
      if (s) restoreProp(s, ch.prop, side === 'before' ? ch.before : ch.after);
    }
    for (const ch of tx.cells.values()) {
      const s = wb.sheetById(ch.sheet);
      if (s) s.put(ch.key, cloneCell(side === 'before' ? ch.before : ch.after));
    }
    const active = side === 'before' ? tx.activeBefore : tx.activeAfter!;
    wb.activeSheet = Math.min(Math.max(0, active), wb.sheets.length - 1);
    const sel = side === 'before' ? tx.selBefore : tx.selAfter!;
    const idx = wb.sheets.findIndex((s) => s.id === sel.sheet);
    if (idx >= 0) {
      wb.activeSheet = idx;
      this.sel = structuredClone(sel.sel);
    }
    this.syncEngine(tx);
    this.emit('edit');
  }

  /* ------------------------------------------------------------ values */

  value(sheet: Sheet, r: number, c: number): Scalar | undefined {
    return this.engine.displayValue(sheet.id, r, c);
  }

  styleOf(sheet: Sheet, r: number, c: number): CellStyle {
    return this.wb.style(this.wb.cellStyleId(sheet, r, c));
  }

  /** Text shown in a cell (number formats applied). */
  displayText(sheet: Sheet, r: number, c: number): { text: string; color?: string; value: Scalar | undefined } {
    const v = this.value(sheet, r, c);
    if (v === undefined || v === null) return { text: '', value: v };
    const fmt = this.styleOf(sheet, r, c).numFmt;
    if (isErr(v)) return { text: asErr(v).error, value: v };
    const f = formatValue(v, fmt);
    return { text: f.text, color: f.color, value: v };
  }

  /** What the formula bar shows for a cell. */
  editText(sheet: Sheet, r: number, c: number): string {
    const cell = sheet.get(r, c);
    if (!cell) return '';
    if (cell.f !== undefined) return '=' + cell.f;
    const v = cell.v;
    if (v === undefined || v === null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (isErr(v)) return asErr(v).error;
    if (typeof v === 'number') {
      const fmt = this.wb.style(cell.s).numFmt;
      if (fmt && isDateFormat(fmt)) {
        const hasDate = /[dy]|m(?!m?:)/i.test(fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, ''));
        const hasTime = /[hs]/i.test(fmt.replace(/"[^"]*"/g, ''));
        const dateFmt = hasDate ? SHORT_DATE : '';
        const timeFmt = hasTime ? (/am\/pm/i.test(fmt) ? 'h:mm:ss AM/PM' : 'h:mm:ss') : '';
        return formatValue(v, [dateFmt, timeFmt].filter(Boolean).join(' ') || SHORT_DATE).text;
      }
      if (fmt && /%/.test(fmt.replace(/"[^"]*"/g, ''))) return `${parseFloat((v * 100).toPrecision(15))}%`;
      return String(parseFloat(v.toPrecision(15)));
    }
    const s = String(v);
    // text that would be read as a number/date/formula keeps a leading apostrophe
    if (s.startsWith('=') || (s.trim() !== '' && parseInput(s).value !== s)) return "'" + s;
    return s;
  }

  /* ------------------------------------------------------------- input */

  /**
   * Converts typed text into a cell. Throws FormulaError for invalid formulas.
   * Returns the new cell and an optional number format to apply.
   */
  interpret(raw: string, sheet: Sheet): { cell: Partial<Cell>; numFmt?: string } {
    if (raw.startsWith('=') && raw.length > 1) {
      const f = normalizeFormula(raw.slice(1));
      const ast = parseFormula(f);
      let numFmt: string | undefined;
      if (ast.t === 'func' && DATE_FNS[ast.name]) numFmt = DATE_FNS[ast.name];
      else {
        // inherit the number format of the first referenced cell (=A1+7 keeps A1's date format)
        const first = tokenize(f).find((t) => t.type === 'ref');
        const plain = !tokenize(f).some((t) => t.type === 'func' && !/^(SUM|MIN|MAX|AVERAGE|ROUND|ABS)$/i.test(t.text));
        if (first && plain) {
          const target = first.sheet ? this.wb.sheetByName(first.sheet) : sheet;
          const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(first.refText!);
          if (target && m) {
            const rr = Number(m[2]) - 1;
            const cc = colIdx(m[1]);
            const fmt = this.styleOf(target, rr, cc).numFmt;
            if (fmt && fmt !== 'General') numFmt = fmt;
          }
        }
      }
      return { cell: { f, v: undefined }, numFmt };
    }
    const p = parseInput(raw);
    return { cell: { v: p.value === null ? undefined : p.value, f: undefined }, numFmt: p.format };
  }

  /** Writes typed input to a cell (or every cell of the selection with `all`). */
  enter(raw: string, opts: { all?: boolean; array?: boolean } = {}): void {
    const sheet = this.sheet;
    const { active } = this.sel;
    const targets: CellPos[] = [];
    if (opts.all) {
      for (const rg of this.sel.ranges) {
        const b = this.clampToUsed(sheet, rg, 5000);
        for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) targets.push({ r, c });
      }
    } else targets.push(active);
    // validate once so errors surface before anything changes
    const base = this.interpret(raw, sheet);
    this.transact('Typing', (tx) => {
      for (const t of targets) {
        const key = cellKey(t.r, t.c);
        const cur = sheet.cells.get(key);
        let cell: Partial<Cell> = base.cell;
        if (base.cell.f !== undefined && (t.r !== active.r || t.c !== active.c)) cell = { f: translateFormula(base.cell.f, t.r - active.r, t.c - active.c), v: undefined };
        const next: Cell = { ...(cur ?? {}), ...cell };
        if (next.f === undefined) delete next.f;
        if (next.v === undefined) delete next.v;
        if (base.numFmt) {
          const curFmt = this.wb.style(cur?.s ?? this.wb.cellStyleId(sheet, t.r, t.c)).numFmt;
          if (!curFmt || curFmt === 'General' || (isDateFormat(base.numFmt) && !isDateFormat(curFmt) && typeof next.v === 'number')) {
            next.s = this.wb.patchStyle(cur?.s ?? this.wb.cellStyleId(sheet, t.r, t.c), { numFmt: base.numFmt });
          }
        } else if (cur?.s === undefined) {
          const inherited = this.wb.cellStyleId(sheet, t.r, t.c);
          if (inherited) next.s = inherited;
        }
        // URLs typed into a cell become links
        if (typeof next.v === 'string' && /^(https?:\/\/|www\.)\S+$/i.test(next.v) && !next.link) next.link = next.v.startsWith('www.') ? `https://${next.v}` : next.v;
        tx.set(sheet, key, next);
      }
    });
  }

  /* ------------------------------------------------------------- clear */

  clampToUsed(sheet: Sheet, rg: Range, pad = 0): Range {
    const ext = this.engine.extent(sheet.id);
    return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(rg.r1, ext.rows + pad)), c2: Math.min(rg.c2, Math.max(rg.c1, ext.cols + pad)) };
  }

  /** Existing cells inside ranges (sparse iteration). */
  cellsIn(sheet: Sheet, ranges: Range[], cb: (key: number, cell: Cell, r: number, c: number) => void): void {
    const total = ranges.reduce((n, rg) => n + (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1), 0);
    if (total > sheet.cells.size * 2) {
      for (const [k, cell] of [...sheet.cells.entries()]) {
        const r = keyRow(k);
        const c = keyCol(k);
        if (ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2)) cb(k, cell, r, c);
      }
      return;
    }
    const seen = new Set<number>();
    for (const rg of ranges)
      for (let r = rg.r1; r <= rg.r2; r++)
        for (let c = rg.c1; c <= rg.c2; c++) {
          const k = cellKey(r, c);
          if (seen.has(k)) continue;
          seen.add(k);
          const cell = sheet.cells.get(k);
          if (cell) cb(k, cell, r, c);
        }
  }

  clear(what: 'all' | 'contents' | 'formats' | 'notes' | 'links'): void {
    const sheet = this.sheet;
    const ranges = this.sel.ranges;
    this.transact(what === 'contents' ? 'Clear contents' : `Clear ${what}`, (tx) => {
      this.cellsIn(sheet, ranges, (k, cell) => {
        if (what === 'all') tx.set(sheet, k, undefined);
        else if (what === 'contents') {
          if (cell.v !== undefined || cell.f !== undefined) tx.patch(sheet, k, { v: undefined, f: undefined });
        } else if (what === 'formats') {
          if (cell.s) tx.patch(sheet, k, { s: undefined });
        } else if (what === 'notes') {
          if (cell.note) tx.patch(sheet, k, { note: undefined });
        } else if (cell.link) tx.patch(sheet, k, { link: undefined });
      });
      if (what === 'all' || what === 'formats') {
        for (const rg of ranges) {
          if (rg.r1 === 0 && rg.r2 === MAX_ROWS - 1) {
            tx.prop(sheet, 'cols');
            for (let c = rg.c1; c <= rg.c2; c++) {
              const ci = sheet.cols.get(c);
              if (ci?.s) sheet.cols.set(c, { ...ci, s: undefined });
            }
          }
          if (rg.c1 === 0 && rg.c2 === MAX_COLS - 1) {
            tx.prop(sheet, 'rows');
            for (let r = rg.r1; r <= rg.r2; r++) {
              const ri = sheet.rows.get(r);
              if (ri?.s) sheet.rows.set(r, { ...ri, s: undefined });
            }
          }
        }
        if (what === 'all') {
          const before = sheet.merges.length;
          const kept = sheet.merges.filter((m) => !ranges.some((rg) => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2));
          if (kept.length !== before) {
            tx.prop(sheet, 'merges');
            sheet.merges = kept;
          }
        }
      }
    });
  }

  /* ------------------------------------------------------------ styles */

  /** Applies a style patch to the selection (whole rows/columns use row/column styles). */
  applyStyle(patch: Partial<CellStyle>, label = 'Format cells', ranges = this.sel.ranges): void {
    const sheet = this.sheet;
    const wb = this.wb;
    this.transact(label, (tx) => {
      for (const rg of ranges) {
        const fullCols = rg.r1 === 0 && rg.r2 === MAX_ROWS - 1;
        const fullRows = rg.c1 === 0 && rg.c2 === MAX_COLS - 1;
        if (fullCols || fullRows) {
          if (fullCols) {
            tx.prop(sheet, 'cols');
            const c2 = fullRows ? Math.max(rg.c1, Math.min(rg.c2, this.engine.extent(sheet.id).cols + 30)) : rg.c2;
            for (let c = rg.c1; c <= c2; c++) {
              const ci = sheet.cols.get(c) ?? {};
              sheet.cols.set(c, { ...ci, s: wb.patchStyle(ci.s, patch) });
            }
          }
          if (fullRows && !fullCols) {
            tx.prop(sheet, 'rows');
            for (let r = rg.r1; r <= rg.r2; r++) {
              const ri = sheet.rows.get(r) ?? {};
              sheet.rows.set(r, { ...ri, s: wb.patchStyle(ri.s, patch) });
            }
          }
          // existing cells inside get the patch too
          this.cellsIn(sheet, [rg], (k, cell) => tx.patch(sheet, k, { s: wb.patchStyle(cell.s, patch) }));
          continue;
        }
        for (let r = rg.r1; r <= rg.r2; r++)
          for (let c = rg.c1; c <= rg.c2; c++) {
            const k = cellKey(r, c);
            const cur = sheet.cells.get(k);
            const base = cur?.s ?? wb.cellStyleId(sheet, r, c);
            const s = wb.patchStyle(base, patch);
            if (s !== (cur?.s ?? 0)) tx.patch(sheet, k, { s: s || undefined });
          }
      }
    });
  }

  /** Style of the active cell (for the ribbon state). */
  activeStyle(): CellStyle {
    return this.styleOf(this.sheet, this.sel.active.r, this.sel.active.c);
  }

  applyBorders(kind: BorderKind, style: BorderStyle = 'thin', color?: string): void {
    const sheet = this.sheet;
    const wb = this.wb;
    const b: Border = { style, ...(color ? { color } : {}) };
    const thick: Border = { style: 'medium', ...(color ? { color } : {}) };
    const dbl: Border = { style: 'double', ...(color ? { color } : {}) };
    this.transact('Borders', (tx) => {
      for (const rg0 of this.sel.ranges) {
        const rg = this.clampToUsed(sheet, rg0, 50);
        for (let r = rg.r1; r <= rg.r2; r++)
          for (let c = rg.c1; c <= rg.c2; c++) {
            const top = r === rg.r1;
            const bottom = r === rg.r2;
            const left = c === rg.c1;
            const right = c === rg.c2;
            const p: Partial<CellStyle> = {};
            switch (kind) {
              case 'none':
                Object.assign(p, { bt: undefined, bb: undefined, bl: undefined, br: undefined });
                break;
              case 'all':
                Object.assign(p, { bt: b, bb: b, bl: b, br: b });
                break;
              case 'outside':
              case 'thickOutside': {
                const x = kind === 'outside' ? b : thick;
                if (top) p.bt = x;
                if (bottom) p.bb = x;
                if (left) p.bl = x;
                if (right) p.br = x;
                break;
              }
              case 'inside':
                if (!top) p.bt = b;
                if (!bottom) p.bb = b;
                if (!left) p.bl = b;
                if (!right) p.br = b;
                break;
              case 'insideH':
                if (!top) p.bt = b;
                if (!bottom) p.bb = b;
                break;
              case 'insideV':
                if (!left) p.bl = b;
                if (!right) p.br = b;
                break;
              case 'top':
                if (top) p.bt = b;
                break;
              case 'bottom':
                if (bottom) p.bb = b;
                break;
              case 'left':
                if (left) p.bl = b;
                break;
              case 'right':
                if (right) p.br = b;
                break;
              case 'topBottom':
                if (top) p.bt = b;
                if (bottom) p.bb = b;
                break;
              case 'thickBottom':
                if (bottom) p.bb = thick;
                break;
              case 'doubleBottom':
                if (bottom) p.bb = dbl;
                break;
              case 'topThickBottom':
                if (top) p.bt = b;
                if (bottom) p.bb = thick;
                break;
              case 'topDoubleBottom':
                if (top) p.bt = b;
                if (bottom) p.bb = dbl;
                break;
            }
            if (!Object.keys(p).length) continue;
            const k = cellKey(r, c);
            const cur = sheet.cells.get(k);
            const s = wb.patchStyle(cur?.s ?? wb.cellStyleId(sheet, r, c), p);
            if (s !== (cur?.s ?? 0)) tx.patch(sheet, k, { s: s || undefined });
          }
      }
    });
  }

  /* ------------------------------------------------------------- merge */

  merge(mode: 'merge' | 'center' | 'across' | 'unmerge'): string | null {
    const sheet = this.sheet;
    const ranges = this.sel.ranges;
    let lost = false;
    const ok = this.transact(mode === 'unmerge' ? 'Unmerge cells' : 'Merge cells', (tx) => {
      tx.prop(sheet, 'merges');
      // remove merges touching the selection
      sheet.merges = sheet.merges.filter((m) => !ranges.some((rg) => m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1));
      if (mode === 'unmerge') return;
      for (const rg of ranges) {
        const rows = mode === 'across' ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => rg.r1 + i) : [rg.r1];
        for (const r of rows) {
          const m: Range = mode === 'across' ? { r1: r, c1: rg.c1, r2: r, c2: rg.c2 } : rg;
          if (m.r1 === m.r2 && m.c1 === m.c2) continue;
          sheet.merges.push(m);
          this.cellsIn(sheet, [m], (k, cell, rr, cc) => {
            if (rr === m.r1 && cc === m.c1) return;
            if (cell.v !== undefined || cell.f !== undefined) {
              lost = true;
              tx.patch(sheet, k, { v: undefined, f: undefined });
            }
          });
        }
      }
      if (mode === 'center') {
        for (const rg of ranges) {
          const k = cellKey(rg.r1, rg.c1);
          const cur = sheet.cells.get(k);
          tx.patch(sheet, k, { s: this.wb.patchStyle(cur?.s ?? this.wb.cellStyleId(sheet, rg.r1, rg.c1), { hAlign: 'center', vAlign: 'middle' }) });
        }
      }
    });
    if (!ok) return null;
    return lost ? 'Merging kept only the upper-left value.' : null;
  }

  /* ------------------------------------------------------------ helpers */

  /** Number of selected cells (whole rows/columns count fully). */
  selectionSize(): number {
    return this.sel.ranges.reduce((n, rg) => n + (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1), 0);
  }

  /** Quick statistics for the status bar. */
  selectionStats(): { count: number; numCount: number; sum: number; avg: number | null; min: number | null; max: number | null } {
    let count = 0;
    let numCount = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    const sheet = this.sheet;
    const seen = new Set<number>();
    for (const rg0 of this.sel.ranges) {
      const rg = this.clampToUsed(sheet, rg0);
      const visit = (r: number, c: number) => {
        const k = cellKey(r, c);
        if (seen.has(k)) return;
        seen.add(k);
        const v = this.value(sheet, r, c);
        if (v === undefined || v === null || v === '') return;
        count++;
        if (typeof v === 'number') {
          numCount++;
          sum += v;
          if (v < min) min = v;
          if (v > max) max = v;
        }
      };
      const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
      if (area > 200000) {
        this.cellsIn(sheet, [rg], (_k, _cell, r, c) => visit(r, c));
      } else for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) visit(r, c);
    }
    return { count, numCount, sum, avg: numCount ? sum / numCount : null, min: numCount ? min : null, max: numCount ? max : null };
  }
}

export type BorderKind =
  | 'none'
  | 'all'
  | 'outside'
  | 'thickOutside'
  | 'inside'
  | 'insideH'
  | 'insideV'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'topBottom'
  | 'thickBottom'
  | 'doubleBottom'
  | 'topThickBottom'
  | 'topDoubleBottom';

function colIdx(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Upper-cases function names/references and closes missing parentheses like Excel. */
export function normalizeFormula(f: string): string {
  const toks = tokenize(f);
  let out = '';
  let depth = 0;
  for (const t of toks) {
    if (t.type === 'func' || t.type === 'bool') out += t.text.toUpperCase();
    else if ((t.type === 'ref' || t.type === 'colrange' || t.type === 'rowrange') && t.refText) out += t.text.slice(0, t.text.length - t.refText.length) + t.refText.toUpperCase();
    else out += t.text;
    if (t.type === 'lparen') depth++;
    if (t.type === 'rparen') depth--;
  }
  if (depth > 0) {
    // only auto-close when that makes the formula valid
    const closed = out + ')'.repeat(depth);
    try {
      parseFormula(closed);
      return closed;
    } catch {
      return out;
    }
  }
  return out;
}

export { FormulaError };
