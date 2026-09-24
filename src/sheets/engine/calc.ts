import './functions';
import { collectFunctions, collectRefs, FormulaError, parseFormula, type Node } from './parser';
import { evaluate, finalize, type EvalContext } from './evaluator';
import { getFunction } from './functions/registry';
import { ERR, FErr, isErr, asErr, isMatrix, type Matrix, type RefValue, type Scalar } from './values';
import { offsetAst } from './refshift';
import { dateToSerial } from '../format/numfmt';
import { MAX_COLS, MAX_ROWS } from '../model/address';
import type { Sheet, Workbook } from '../model/workbook';

/* ================================================================ keys */

const STRIDE = 2 ** 34;
export const gkey = (sheet: number, r: number, c: number) => sheet * STRIDE + r * MAX_COLS + c;
export const gSheet = (g: number) => Math.floor(g / STRIDE);
export const gRow = (g: number) => Math.floor((g % STRIDE) / MAX_COLS);
export const gCol = (g: number) => (g % STRIDE) % MAX_COLS;

interface Area {
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

interface RangeEntry extends Area {
  dep: number;
}

/* ======================================================= range index */

const BLOCK = 32;
const MAX_BLOCKS = 8;

/** Finds formulas that depend on ranges containing a given cell. */
class RangeIndex {
  private buckets = new Map<number, Set<RangeEntry>>();
  private wide = new Map<number, Set<RangeEntry>>();

  private bucketKeys(e: RangeEntry): number[] | null {
    const b1 = Math.floor(e.c1 / BLOCK);
    const b2 = Math.floor(e.c2 / BLOCK);
    if (b2 - b1 + 1 > MAX_BLOCKS) return null;
    const out: number[] = [];
    for (let b = b1; b <= b2; b++) out.push(e.sheet * 1024 + b);
    return out;
  }

  add(e: RangeEntry): void {
    const keys = this.bucketKeys(e);
    if (!keys) {
      let set = this.wide.get(e.sheet);
      if (!set) this.wide.set(e.sheet, (set = new Set()));
      set.add(e);
      return;
    }
    for (const k of keys) {
      let set = this.buckets.get(k);
      if (!set) this.buckets.set(k, (set = new Set()));
      set.add(e);
    }
  }

  remove(e: RangeEntry): void {
    const keys = this.bucketKeys(e);
    if (!keys) {
      this.wide.get(e.sheet)?.delete(e);
      return;
    }
    for (const k of keys) this.buckets.get(k)?.delete(e);
  }

  query(sheet: number, r: number, c: number, cb: (dep: number) => void): void {
    const scan = (set: Set<RangeEntry> | undefined) => {
      if (!set) return;
      for (const e of set) if (r >= e.r1 && r <= e.r2 && c >= e.c1 && c <= e.c2) cb(e.dep);
    };
    scan(this.buckets.get(sheet * 1024 + Math.floor(c / BLOCK)));
    scan(this.wide.get(sheet));
  }

  clear(): void {
    this.buckets.clear();
    this.wide.clear();
  }
}

/* ========================================================== contexts */

class Ctx implements EvalContext {
  constructor(
    private eng: CalcEngine,
    public sheet: number,
    public row: number,
    public col: number,
  ) {}
  cell(s: number, r: number, c: number): Scalar {
    return this.eng.valueAt(s, r, c);
  }
  each(ref: RefValue, cb: (v: Scalar, r: number, c: number) => void): void {
    this.eng.eachIn(ref, cb);
  }
  extent(s: number) {
    return this.eng.extent(s);
  }
  sheetId(name: string) {
    return this.eng.wb.sheetByName(name)?.id;
  }
  sheetName(id: number) {
    return this.eng.sheet(id)?.name;
  }
  sheetIds() {
    return this.eng.wb.sheets.map((s) => s.id);
  }
  definedName(name: string, sheet: number) {
    return this.eng.nameNode(name, sheet);
  }
  formulaText(s: number, r: number, c: number) {
    return this.eng.sheet(s)?.get(r, c)?.f;
  }
  spillRange(s: number, r: number, c: number) {
    return this.eng.spillRef(s, r, c);
  }
  cellFormat(s: number, r: number, c: number) {
    const sh = this.eng.sheet(s);
    if (!sh) return undefined;
    return this.eng.wb.style(this.eng.wb.cellStyleId(sh, r, c)).numFmt;
  }
  rowState(s: number, r: number): 'hidden' | 'filtered' | undefined {
    if (this.eng.filteredRows.get(s)?.has(r)) return 'filtered';
    return this.eng.sheet(s)?.rows.get(r)?.hidden ? 'hidden' : undefined;
  }
  now() {
    return this.eng.nowSerial();
  }
  random() {
    return this.eng.random();
  }
  markVolatile() {
    this.eng.flagVolatile();
  }
}

/* ============================================================ engine */

export interface CalcStats {
  formulas: number;
  lastRecalcMs: number;
  circular: number[];
}

export class CalcEngine {
  readonly wb: Workbook;
  private sheets = new Map<number, Sheet>();
  private asts = new Map<number, Node | FormulaError>();
  private deps = new Map<number, { cells: number[]; ranges: RangeEntry[] }>();
  private cellIndex = new Map<number, Set<number>>();
  private ranges = new RangeIndex();
  private volatileCells = new Set<number>();
  private dirty = new Set<number>();
  private computing = new Set<number>();
  private spillVals = new Map<number, Map<number, Scalar>>();
  private spillOwner = new Map<number, number>();
  private spillAreas = new Map<number, Area>();
  private blocked = new Map<number, Area>();
  private spillExtent = new Map<number, { rows: number; cols: number }>();
  private names = new Map<string, Node | null>();
  private pending: number[] = [];
  private volatileHit = false;
  private nowValue = 0;
  circular = new Set<number>();
  stats: CalcStats = { formulas: 0, lastRecalcMs: 0, circular: [] };
  random: () => number = Math.random;
  clock: () => Date = () => new Date();
  /** Rows hidden by the auto-filter (maintained by the document layer). */
  filteredRows = new Map<number, Set<number>>();

  constructor(wb: Workbook) {
    this.wb = wb;
  }

  sheet(id: number): Sheet | undefined {
    let s = this.sheets.get(id);
    if (!s) {
      s = this.wb.sheetById(id);
      if (s) this.sheets.set(id, s);
    }
    return s;
  }

  nowSerial(): number {
    if (!this.nowValue) this.nowValue = dateToSerial(this.clock());
    return this.nowValue;
  }

  flagVolatile(): void {
    this.volatileHit = true;
  }

  /* ------------------------------------------------------ registration */

  /** Re-parses every formula and recalculates the whole workbook. */
  rebuild(): void {
    this.sheets.clear();
    for (const s of this.wb.sheets) this.sheets.set(s.id, s);
    this.asts.clear();
    this.deps.clear();
    this.cellIndex.clear();
    this.ranges.clear();
    this.volatileCells.clear();
    this.dirty.clear();
    this.spillVals.clear();
    this.spillOwner.clear();
    this.spillAreas.clear();
    this.blocked.clear();
    this.spillExtent.clear();
    this.names.clear();
    this.circular.clear();
    const roots: number[] = [];
    for (const s of this.wb.sheets) {
      for (const [k, cell] of s.cells) {
        if (cell.f === undefined) continue;
        const g = s.id * STRIDE + k;
        this.register(g, cell.f);
        roots.push(g);
      }
    }
    for (const g of roots) this.dirty.add(g);
    this.recalc(roots);
  }

  /** Called after cells were edited: re-registers formulas and recalculates dependents. */
  cellsChanged(changes: { sheet: number; key: number }[]): void {
    const roots: number[] = [];
    for (const { sheet, key } of changes) {
      const g = sheet * STRIDE + key;
      const s = this.sheet(sheet);
      const cell = s?.cells.get(key);
      if (this.asts.has(g)) this.unregister(g);
      if (cell?.f !== undefined) {
        this.register(g, cell.f);
        this.dirty.add(g);
      }
      roots.push(g);
      const owner = this.spillOwner.get(g);
      if (owner !== undefined) {
        this.dirty.add(owner);
        roots.push(owner);
      }
      for (const [anchor, area] of this.blocked) {
        if (area.sheet === sheet && inArea(area, Math.floor(key / MAX_COLS), key % MAX_COLS)) {
          this.dirty.add(anchor);
          roots.push(anchor);
        }
      }
    }
    if (this.wb.calcAuto) this.recalc(roots);
  }

  /** Invalidates cached defined names (after the name manager changed them). */
  namesChanged(): void {
    this.rebuild();
  }

  private register(g: number, text: string): void {
    let ast: Node | FormulaError;
    try {
      ast = parseFormula(text);
    } catch (e) {
      ast = e instanceof FormulaError ? e : new FormulaError(String(e));
    }
    this.asts.set(g, ast);
    if (ast instanceof FormulaError) return;
    const sheet = gSheet(g);
    const cells: number[] = [];
    const ranges: RangeEntry[] = [];
    const seenNames = new Set<string>();
    const visit = (node: Node, ctxSheet: number) => {
      for (const ref of collectRefs(node)) {
        if (ref.t === 'ref') {
          const sid = ref.sheet ? this.wb.sheetByName(ref.sheet)?.id : ctxSheet;
          if (sid !== undefined) cells.push(gkey(sid, ref.r, ref.c));
        } else if (ref.t === 'range') {
          const sid = ref.sheet ? this.wb.sheetByName(ref.sheet)?.id : ctxSheet;
          if (sid !== undefined) ranges.push({ sheet: sid, r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2, dep: g });
        } else if (ref.t === 'name') {
          const key = `${ref.sheet ?? ''}!${ref.name}`;
          if (seenNames.has(key)) continue;
          seenNames.add(key);
          const sid = ref.sheet ? this.wb.sheetByName(ref.sheet)?.id ?? ctxSheet : ctxSheet;
          const def = this.nameNode(ref.name, sid);
          if (def) visit(def, sid);
        }
      }
      for (const f of collectFunctions(node)) {
        if (getFunction(f)?.volatile) this.volatileCells.add(g);
      }
    };
    visit(ast, sheet);
    for (const c of cells) {
      let set = this.cellIndex.get(c);
      if (!set) this.cellIndex.set(c, (set = new Set()));
      set.add(g);
    }
    for (const r of ranges) {
      // single-cell ranges behave like cell refs
      if (r.r1 === r.r2 && r.c1 === r.c2) {
        const c = gkey(r.sheet, r.r1, r.c1);
        cells.push(c);
        let set = this.cellIndex.get(c);
        if (!set) this.cellIndex.set(c, (set = new Set()));
        set.add(g);
      } else this.ranges.add(r);
    }
    this.deps.set(g, { cells, ranges: ranges.filter((r) => r.r1 !== r.r2 || r.c1 !== r.c2) });
  }

  private unregister(g: number): void {
    const d = this.deps.get(g);
    if (d) {
      for (const c of d.cells) this.cellIndex.get(c)?.delete(g);
      for (const r of d.ranges) this.ranges.remove(r);
    }
    this.deps.delete(g);
    this.asts.delete(g);
    this.volatileCells.delete(g);
    this.dirty.delete(g);
    this.circular.delete(g);
    this.clearSpill(g);
  }

  /* ---------------------------------------------------------- recalc */

  /** Recalculates everything that depends on the given cells (plus volatile cells). */
  recalc(roots: number[] = []): void {
    const t0 = performance.now();
    this.nowValue = 0;
    let seeds = [...roots, ...this.volatileCells];
    for (const g of this.volatileCells) this.dirty.add(g);
    let rounds = 0;
    while (seeds.length && rounds++ < 20) {
      const order = this.topoOrder(seeds);
      this.pending = [];
      for (const g of order) if (this.dirty.has(g)) this.compute(g);
      seeds = this.pending;
    }
    // anything still dirty (e.g. unreachable due to missing sheets) is computed directly
    for (const g of [...this.dirty]) this.compute(g);
    this.stats = { formulas: this.asts.size, lastRecalcMs: performance.now() - t0, circular: [...this.circular] };
  }

  /** Full recalculation (F9 / manual mode). */
  recalcAll(): void {
    const all = [...this.asts.keys()];
    for (const g of all) this.dirty.add(g);
    this.recalc(all);
  }

  private dependents(g: number): number[] {
    const out: number[] = [];
    const set = this.cellIndex.get(g);
    if (set) for (const d of set) out.push(d);
    this.ranges.query(gSheet(g), gRow(g), gCol(g), (d) => out.push(d));
    return out;
  }

  /** Reverse post-order DFS over dependents: precedents come before dependents. */
  private topoOrder(seeds: number[]): number[] {
    const visited = new Set<number>();
    const post: number[] = [];
    const stackNode: number[] = [];
    const stackDeps: number[][] = [];
    const stackIdx: number[] = [];
    for (const s of seeds) {
      if (visited.has(s)) continue;
      visited.add(s);
      stackNode.push(s);
      stackDeps.push(this.dependents(s));
      stackIdx.push(0);
      while (stackNode.length) {
        const top = stackNode.length - 1;
        const deps = stackDeps[top];
        if (stackIdx[top] < deps.length) {
          const d = deps[stackIdx[top]++];
          if (!visited.has(d)) {
            visited.add(d);
            stackNode.push(d);
            stackDeps.push(this.dependents(d));
            stackIdx.push(0);
          }
        } else {
          post.push(stackNode.pop()!);
          stackDeps.pop();
          stackIdx.pop();
        }
      }
    }
    post.reverse();
    for (const g of post) if (this.asts.has(g)) this.dirty.add(g);
    return post;
  }

  private compute(g: number): void {
    if (!this.dirty.has(g)) return;
    const sid = gSheet(g);
    const r = gRow(g);
    const c = gCol(g);
    const sheet = this.sheet(sid);
    const cell = sheet?.cells.get(r * MAX_COLS + c);
    const ast = this.asts.get(g);
    if (!sheet || !cell || cell.f === undefined || !ast) {
      this.dirty.delete(g);
      return;
    }
    this.computing.add(g);
    let result: Scalar | Matrix;
    const prevVolatile = this.volatileHit;
    this.volatileHit = false;
    if (ast instanceof FormulaError) result = ERR.NAME;
    else {
      const ctx = new Ctx(this, sid, r, c);
      try {
        result = finalize(evaluate(ast, ctx), ctx);
      } catch (e) {
        result = e instanceof FErr ? e : isErr(e) ? asErr(e) : ERR.NUM;
      }
    }
    if (this.volatileHit) this.volatileCells.add(g);
    this.volatileHit = prevVolatile;
    this.computing.delete(g);
    this.dirty.delete(g);
    let value: Scalar;
    if (isMatrix(result)) value = this.applySpill(g, sheet, r, c, result);
    else {
      if (this.spillAreas.has(g) || this.blocked.has(g)) this.clearSpill(g);
      value = result;
    }
    if (this.circular.has(g) && !isErr(value)) this.circular.delete(g);
    cell.v = value;
  }

  /* ----------------------------------------------------------- values */

  valueAt(s: number, r: number, c: number): Scalar {
    const sheet = this.sheet(s);
    if (!sheet) return ERR.REF;
    const k = r * MAX_COLS + c;
    const cell = sheet.cells.get(k);
    if (cell) {
      if (cell.f !== undefined) {
        const g = s * STRIDE + k;
        if (this.dirty.has(g)) {
          if (this.computing.has(g)) {
            this.circular.add(g);
            return ERR.CIRC;
          }
          this.compute(g);
        }
        return cell.v ?? null;
      }
      if (cell.v !== undefined) return cell.v;
    }
    const sp = this.spillVals.get(s);
    if (sp) {
      const v = sp.get(k);
      if (v !== undefined) {
        const owner = this.spillOwner.get(s * STRIDE + k);
        if (owner !== undefined && this.dirty.has(owner)) {
          if (this.computing.has(owner)) {
            this.circular.add(owner);
            return ERR.CIRC;
          }
          this.compute(owner);
          return this.spillVals.get(s)?.get(k) ?? null;
        }
        return v;
      }
    }
    return null;
  }

  /** Displayed value of a cell (no evaluation side effects). */
  displayValue(s: number, r: number, c: number): Scalar | undefined {
    const sheet = this.sheet(s);
    if (!sheet) return undefined;
    const k = r * MAX_COLS + c;
    const cell = sheet.cells.get(k);
    if (cell && (cell.f !== undefined || cell.v !== undefined)) return cell.v ?? null;
    return this.spillVals.get(s)?.get(k);
  }

  /** Anchor of the spill covering a cell (for the UI's spill border / greyed formula). */
  spillAnchorOf(s: number, r: number, c: number): { r: number; c: number } | undefined {
    const owner = this.spillOwner.get(gkey(s, r, c));
    if (owner === undefined) return undefined;
    return { r: gRow(owner), c: gCol(owner) };
  }

  spillAreaOf(s: number, r: number, c: number): Area | undefined {
    return this.spillAreas.get(gkey(s, r, c)) ?? this.blocked.get(gkey(s, r, c));
  }

  isSpillBlocked(s: number, r: number, c: number): boolean {
    return this.blocked.has(gkey(s, r, c));
  }

  eachIn(ref: RefValue, cb: (v: Scalar, r: number, c: number) => void): void {
    const sheet = this.sheet(ref.sheet);
    if (!sheet) throw ERR.REF;
    const ext = this.extent(ref.sheet);
    const r2 = Math.min(ref.r2, ext.rows);
    const c2 = Math.min(ref.c2, ext.cols);
    if (r2 < ref.r1 || c2 < ref.c1) return;
    const area = (r2 - ref.r1 + 1) * (c2 - ref.c1 + 1);
    const sp = this.spillVals.get(ref.sheet);
    const population = sheet.cells.size + (sp?.size ?? 0);
    if (area <= population * 2) {
      for (let r = ref.r1; r <= r2; r++)
        for (let c = ref.c1; c <= c2; c++) {
          const v = this.valueAt(ref.sheet, r, c);
          if (v !== null) cb(v, r, c);
        }
      return;
    }
    const keys: number[] = [];
    for (const k of sheet.cells.keys()) {
      const r = Math.floor(k / MAX_COLS);
      if (r < ref.r1 || r > r2) continue;
      const c = k % MAX_COLS;
      if (c >= ref.c1 && c <= c2) keys.push(k);
    }
    if (sp)
      for (const k of sp.keys()) {
        if (sheet.cells.get(k)?.v !== undefined) continue;
        const r = Math.floor(k / MAX_COLS);
        if (r < ref.r1 || r > r2) continue;
        const c = k % MAX_COLS;
        if (c >= ref.c1 && c <= c2) keys.push(k);
      }
    for (const k of keys) {
      const r = Math.floor(k / MAX_COLS);
      const c = k % MAX_COLS;
      const v = this.valueAt(ref.sheet, r, c);
      if (v !== null) cb(v, r, c);
    }
  }

  extent(s: number): { rows: number; cols: number } {
    const sheet = this.sheet(s);
    if (!sheet) return { rows: -1, cols: -1 };
    const e = sheet.extent();
    let se = this.spillExtent.get(s);
    if (!se) {
      se = { rows: -1, cols: -1 };
      for (const a of this.spillAreas.values()) {
        if (a.sheet !== s) continue;
        if (a.r2 > se.rows) se.rows = a.r2;
        if (a.c2 > se.cols) se.cols = a.c2;
      }
      this.spillExtent.set(s, se);
    }
    return { rows: Math.max(e.rows, se.rows), cols: Math.max(e.cols, se.cols) };
  }

  /* ----------------------------------------------------------- spills */

  private applySpill(g: number, sheet: Sheet, r: number, c: number, m: Matrix): Scalar {
    const R = m.data.length;
    const C = m.data[0].length;
    const area: Area = { sheet: sheet.id, r1: r, c1: c, r2: r + R - 1, c2: c + C - 1 };
    let blocked = area.r2 >= MAX_ROWS || area.c2 >= MAX_COLS;
    if (!blocked) {
      outer: for (let i = 0; i < R; i++)
        for (let j = 0; j < C; j++) {
          if (i === 0 && j === 0) continue;
          const k = (r + i) * MAX_COLS + (c + j);
          const cell = sheet.cells.get(k);
          if (cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null))) {
            blocked = true;
            break outer;
          }
          const owner = this.spillOwner.get(sheet.id * STRIDE + k);
          if (owner !== undefined && owner !== g) {
            blocked = true;
            break outer;
          }
        }
      if (!blocked) for (const mg of sheet.merges) if (mg.r1 <= area.r2 && mg.r2 >= area.r1 && mg.c1 <= area.c2 && mg.c2 >= area.c1) blocked = true;
    }
    if (blocked) {
      this.clearSpill(g);
      this.blocked.set(g, area);
      return ERR.SPILL;
    }
    this.blocked.delete(g);
    let sp = this.spillVals.get(sheet.id);
    if (!sp) this.spillVals.set(sheet.id, (sp = new Map()));
    const old = this.spillAreas.get(g);
    if (old) {
      for (let rr = old.r1; rr <= old.r2; rr++)
        for (let cc = old.c1; cc <= old.c2; cc++) {
          if ((rr === r && cc === c) || inArea(area, rr, cc)) continue;
          const k = rr * MAX_COLS + cc;
          sp.delete(k);
          this.spillOwner.delete(sheet.id * STRIDE + k);
          this.pending.push(sheet.id * STRIDE + k);
        }
    }
    for (let i = 0; i < R; i++)
      for (let j = 0; j < C; j++) {
        if (i === 0 && j === 0) continue;
        const k = (r + i) * MAX_COLS + (c + j);
        const gk = sheet.id * STRIDE + k;
        const v = m.data[i][j];
        const prev = sp.get(k);
        if (!sameScalar(prev, v)) this.pending.push(gk);
        sp.set(k, v);
        this.spillOwner.set(gk, g);
      }
    this.spillAreas.set(g, area);
    this.spillExtent.delete(sheet.id);
    return m.data[0][0];
  }

  private clearSpill(g: number): void {
    const area = this.spillAreas.get(g);
    this.blocked.delete(g);
    if (!area) return;
    const sp = this.spillVals.get(area.sheet);
    for (let r = area.r1; r <= area.r2; r++)
      for (let c = area.c1; c <= area.c2; c++) {
        const k = r * MAX_COLS + c;
        const gk = area.sheet * STRIDE + k;
        if (gk === g) continue;
        sp?.delete(k);
        this.spillOwner.delete(gk);
        this.pending.push(gk);
      }
    this.spillAreas.delete(g);
    this.spillExtent.delete(area.sheet);
  }

  spillRef(s: number, r: number, c: number): RefValue | undefined {
    const g = gkey(s, r, c);
    if (this.dirty.has(g) && !this.computing.has(g)) this.compute(g);
    const a = this.spillAreas.get(g);
    if (a) return { kind: 'ref', sheet: a.sheet, r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 };
    const cell = this.sheet(s)?.cells.get(r * MAX_COLS + c);
    if (cell?.f !== undefined && !isErr(cell.v)) return { kind: 'ref', sheet: s, r1: r, c1: c, r2: r, c2: c };
    return undefined;
  }

  /* ------------------------------------------------------------ names */

  nameNode(name: string, sheet: number): Node | undefined {
    const key = `${sheet}|${name}`;
    if (this.names.has(key)) return this.names.get(key) ?? undefined;
    const upper = name.toUpperCase();
    const def =
      this.wb.names.find((n) => n.sheet === sheet && n.name.toUpperCase() === upper) ?? this.wb.names.find((n) => n.sheet === undefined && n.name.toUpperCase() === upper);
    let node: Node | null = null;
    if (def) {
      try {
        node = parseFormula(def.ref.replace(/^=/, ''));
      } catch {
        node = { t: 'err', v: '#NAME?' };
      }
    }
    this.names.set(key, node);
    return node ?? undefined;
  }

  /* ------------------------------------------------------- evaluation */

  /**
   * Evaluates formula text in the context of a cell without storing it
   * (conditional formatting, validation rules, the formula bar preview).
   * `anchor` makes relative references relative to another cell (CF ranges).
   */
  evaluateText(text: string, sheet: number, r: number, c: number, anchor?: { r: number; c: number }): Scalar | Matrix {
    let ast: Node;
    try {
      ast = parseFormula(text.replace(/^=/, ''));
    } catch {
      return ERR.NAME;
    }
    return this.evaluateAst(ast, sheet, r, c, anchor);
  }

  evaluateAst(ast: Node, sheet: number, r: number, c: number, anchor?: { r: number; c: number }): Scalar | Matrix {
    const node = anchor ? offsetAst(ast, r - anchor.r, c - anchor.c) : ast;
    const ctx = new Ctx(this, sheet, r, c);
    try {
      return finalize(evaluate(node, ctx), ctx);
    } catch (e) {
      return e instanceof FErr ? e : ERR.VALUE;
    }
  }

  /** Values of a reference as a 2D array (for charts and data tools). */
  rangeValues(sheet: number, r1: number, c1: number, r2: number, c2: number): Scalar[][] {
    const out: Scalar[][] = [];
    for (let r = r1; r <= r2; r++) {
      const row: Scalar[] = [];
      for (let c = c1; c <= c2; c++) row.push(this.valueAt(sheet, r, c));
      out.push(row);
    }
    return out;
  }

  hasFormulas(): boolean {
    return this.asts.size > 0;
  }

  /** Cells the given cell depends on directly (trace precedents). */
  precedentsOf(s: number, r: number, c: number): { cells: { sheet: number; r: number; c: number }[]; ranges: Area[] } {
    const d = this.deps.get(gkey(s, r, c));
    if (!d) return { cells: [], ranges: [] };
    return { cells: d.cells.map((g) => ({ sheet: gSheet(g), r: gRow(g), c: gCol(g) })), ranges: d.ranges.map(({ sheet, r1, c1, r2, c2 }) => ({ sheet, r1, c1, r2, c2 })) };
  }

  /** Formula cells that use the given cell directly (trace dependents). */
  dependentsOf(s: number, r: number, c: number): { sheet: number; r: number; c: number }[] {
    return this.dependents(gkey(s, r, c)).map((g) => ({ sheet: gSheet(g), r: gRow(g), c: gCol(g) }));
  }
}

function inArea(a: Area, r: number, c: number): boolean {
  return r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2;
}

function sameScalar(a: Scalar | undefined, b: Scalar): boolean {
  if (a === b) return true;
  if (isErr(a) && isErr(b)) return asErr(a).error === asErr(b).error;
  return false;
}
