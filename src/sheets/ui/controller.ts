import { FormulaError, tokenize } from '../engine/parser';
import { cycleAbsolute } from '../engine/refshift';
import { addrName, cellKey, colName, MAX_COLS, MAX_ROWS, parseRange, quoteSheet, type Range } from '../model/address';
import { geometryOf, type SheetGeometry } from '../grid/geometry';
import type { CellPos, SheetDoc } from '../doc';

export interface EditState {
  sheet: number;
  r: number;
  c: number;
  text: string;
  caret: number;
  /** 'enter' = started by typing (arrows commit), 'edit' = F2/double-click (arrows move the caret). */
  mode: 'enter' | 'edit';
  /** Where the caret lives. */
  source: 'cell' | 'bar';
  /** Text span of the reference being pointed at with mouse/keys. */
  point?: { start: number; end: number; anchor: CellPos; cursor: CellPos };
  /** Original text (for Escape). */
  original: string;
}

export const REF_COLORS = ['#3e63dd', '#e5484d', '#8e4ec6', '#30a46c', '#f76b15', '#0090ff', '#d6409f', '#12a594'];

type Listener = () => void;

/** UI state shared by the grid, the formula bar, the ribbon and the status bar. */
export class SheetUI {
  edit: EditState | null = null;
  findHits: Set<number> | null = null;
  clipPhase = 0;
  showFormulas = false;
  /** Where a run of Tab presses started, so Enter can return to that column (like Excel). */
  private tabRun: { sheet: number; col: number; at: CellPos } | null = null;
  version = 0;
  /** Scroll offsets (sheet px, unzoomed) per sheet id. */
  scroll = new Map<number, { x: number; y: number }>();
  /** Extra virtual rows/cols so the scrollbars can reach far cells. */
  reach = new Map<number, { rows: number; cols: number }>();
  viewport = { width: 800, height: 600, headerW: 46, headerH: 24 };
  /** Chart or picture selected on the sheet. */
  selectedObject: { kind: 'chart' | 'image'; id: string } | null = null;
  onEditChart?: (id: string) => void;
  onScrollRequest?: () => void;
  onError?: (msg: string, detail?: string) => void;
  onNotify?: (msg: string) => void;
  private listeners = new Set<Listener>();

  constructor(public doc: SheetDoc) {}

  subscribe = (fn: Listener) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };
  getVersion = () => this.version;
  emit(): void {
    this.version++;
    for (const l of this.listeners) l();
  }

  get sheet() {
    return this.doc.sheet;
  }

  get zoom(): number {
    return this.doc.sheet.view.zoom || 1;
  }

  setZoom(z: number): void {
    const v = Math.max(0.25, Math.min(4, Math.round(z * 100) / 100));
    this.doc.sheet.view.zoom = v;
    this.emit();
    this.doc.emit('view');
  }

  geo(): SheetGeometry {
    return geometryOf(this.doc.sheet, this.doc.version);
  }

  getScroll(): { x: number; y: number } {
    const id = this.doc.sheet.id;
    let s = this.scroll.get(id);
    if (!s) {
      const g = this.geo();
      const v = this.doc.sheet.view;
      s = { x: Math.max(0, g.cols.pos(v.scrollCol) - g.cols.pos(v.freezeCols)), y: Math.max(0, g.rows.pos(v.scrollRow) - g.rows.pos(v.freezeRows)) };
      this.scroll.set(id, s);
    }
    return s;
  }

  setScroll(x: number, y: number): void {
    const s = this.getScroll();
    if (s.x === x && s.y === y) return;
    s.x = Math.max(0, x);
    s.y = Math.max(0, y);
    const g = this.geo();
    const v = this.doc.sheet.view;
    v.scrollRow = g.rows.at(g.rows.pos(v.freezeRows) + s.y);
    v.scrollCol = g.cols.at(g.cols.pos(v.freezeCols) + s.x);
  }

  getReach(): { rows: number; cols: number } {
    const id = this.doc.sheet.id;
    let r = this.reach.get(id);
    if (!r) {
      r = { rows: 0, cols: 0 };
      this.reach.set(id, r);
    }
    return r;
  }

  /* ------------------------------------------------------ navigation */

  /** Scrolls so that a cell is visible (called after keyboard moves). */
  ensureVisible(r: number, c: number): void {
    const g = this.geo();
    const z = this.zoom;
    const v = this.doc.sheet.view;
    const s = { ...this.getScroll() };
    const reach = this.getReach();
    if (r + 20 > reach.rows) reach.rows = Math.min(MAX_ROWS, r + 50);
    if (c + 5 > reach.cols) reach.cols = Math.min(MAX_COLS, c + 10);
    const vw = (this.viewport.width - this.viewport.headerW) / z - g.cols.pos(v.freezeCols);
    const vh = (this.viewport.height - this.viewport.headerH) / z - g.rows.pos(v.freezeRows);
    if (r >= v.freezeRows) {
      const top = g.rows.pos(r) - g.rows.pos(v.freezeRows);
      const bottom = top + g.rows.sizeOf(r);
      if (top < s.y) s.y = top;
      else if (bottom > s.y + vh) s.y = Math.max(0, bottom - vh + 2);
    }
    if (c >= v.freezeCols) {
      const left = g.cols.pos(c) - g.cols.pos(v.freezeCols);
      const right = left + g.cols.sizeOf(c);
      if (left < s.x) s.x = left;
      else if (right > s.x + vw) s.x = Math.max(0, Math.min(left, right - vw + 2));
    }
    const cur = this.getScroll();
    if (s.x !== cur.x || s.y !== cur.y) {
      this.setScroll(s.x, s.y);
      this.onScrollRequest?.();
    }
  }

  select(range: Range, active?: CellPos, anchor?: CellPos, add = false): void {
    const doc = this.doc;
    const a = active ?? { r: range.r1, c: range.c1 };
    const ranges = add ? [...doc.sel.ranges, range] : [range];
    doc.sel = { ranges, active: a, anchor: anchor ?? a };
    this.emit();
  }

  selectCell(r: number, c: number, extend = false): void {
    const doc = this.doc;
    const sel = doc.sel;
    if (extend) {
      const an = sel.anchor;
      const rg: Range = { r1: Math.min(an.r, r), c1: Math.min(an.c, c), r2: Math.max(an.r, r), c2: Math.max(an.c, c) };
      doc.sel = { ranges: [...sel.ranges.slice(0, -1), rg], active: sel.active, anchor: an };
    } else {
      const m = doc.sheet.mergeAt(r, c);
      const rg = m ?? { r1: r, c1: c, r2: r, c2: c };
      doc.sel = { ranges: [rg], active: { r: m?.r1 ?? r, c: m?.c1 ?? c }, anchor: { r: m?.r1 ?? r, c: m?.c1 ?? c } };
    }
    this.ensureVisible(r, c);
    this.emit();
  }

  /** Arrow-key movement (skipping hidden rows/cols and merged areas). */
  move(dr: number, dc: number, extend = false, jump = false): void {
    const doc = this.doc;
    const sel = doc.sel;
    const g = this.geo();
    const from = extend ? this.extendCursor() : sel.active;
    let { r, c } = from;
    if (jump) [r, c] = this.jumpTarget(r, c, dr, dc);
    else {
      const m = doc.sheet.mergeAt(r, c);
      if (dr > 0 && m) r = m.r2;
      if (dc > 0 && m) c = m.c2;
      if (dr < 0 && m) r = m.r1;
      if (dc < 0 && m) c = m.c1;
      if (dr) r = g.rows.step(r, dr > 0 ? 1 : -1);
      if (dc) c = g.cols.step(c, dc > 0 ? 1 : -1);
    }
    r = Math.max(0, Math.min(MAX_ROWS - 1, r));
    c = Math.max(0, Math.min(MAX_COLS - 1, c));
    if (extend) {
      const an = sel.anchor;
      let rg: Range = { r1: Math.min(an.r, r), c1: Math.min(an.c, c), r2: Math.max(an.r, r), c2: Math.max(an.c, c) };
      rg = this.withMerges(rg);
      doc.sel = { ranges: [...sel.ranges.slice(0, -1), rg], active: sel.active, anchor: an };
      this.cursor = { r, c };
      this.ensureVisible(r, c);
      this.emit();
      return;
    }
    this.cursor = null;
    this.selectCell(r, c);
  }

  private cursor: CellPos | null = null;

  private extendCursor(): CellPos {
    const sel = this.doc.sel;
    if (this.cursor) return this.cursor;
    const rg = sel.ranges[sel.ranges.length - 1];
    const an = sel.anchor;
    return { r: an.r === rg.r1 ? rg.r2 : rg.r1, c: an.c === rg.c1 ? rg.c2 : rg.c1 };
  }

  withMerges(rg: Range): Range {
    let out = rg;
    for (let guard = 0; guard < 6; guard++) {
      let changed = false;
      for (const m of this.doc.sheet.merges) {
        if (m.r1 <= out.r2 && m.r2 >= out.r1 && m.c1 <= out.c2 && m.c2 >= out.c1) {
          const n = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) };
          if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) {
            out = n;
            changed = true;
          }
        }
      }
      if (!changed) break;
    }
    return out;
  }

  /** Ctrl+Arrow: jump to the edge of the data region. */
  jumpTarget(r: number, c: number, dr: number, dc: number): [number, number] {
    const doc = this.doc;
    const sheet = doc.sheet;
    const filled = (rr: number, cc: number) => {
      const v = doc.value(sheet, rr, cc);
      return v !== undefined && v !== null && v !== '';
    };
    const ext = doc.engine.extent(sheet.id);
    const maxR = dr ? MAX_ROWS - 1 : r;
    const maxC = dc ? MAX_COLS - 1 : c;
    let rr = r;
    let cc = c;
    const next = () => {
      rr += dr;
      cc += dc;
    };
    const inBounds = () => rr >= 0 && cc >= 0 && rr <= maxR && cc <= maxC;
    const cur = filled(r, c);
    next();
    if (!inBounds()) return [r, c];
    if (cur && filled(rr, cc)) {
      // run to the last filled cell
      while (inBounds() && filled(rr, cc)) next();
      return [rr - dr, cc - dc];
    }
    // skip blanks to the next filled cell (or the sheet edge)
    while (inBounds() && !filled(rr, cc)) {
      if ((dr > 0 && rr > ext.rows) || (dc > 0 && cc > ext.cols)) return [dr > 0 ? MAX_ROWS - 1 : rr, dc > 0 ? MAX_COLS - 1 : cc];
      next();
    }
    if (!inBounds()) return [Math.max(0, Math.min(maxR, rr - dr)), Math.max(0, Math.min(maxC, cc - dc))];
    return [rr, cc];
  }

  /** Enter/Tab movement inside a multi-cell selection. */
  cycle(dr: number, dc: number): void {
    const sel = this.doc.sel;
    const rg = sel.ranges[sel.ranges.length - 1];
    const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
    if (!multi) {
      // Enter after a run of Tabs returns to the column where the run started, one row down
      const { r, c } = sel.active;
      const sheet = this.doc.sheet.id;
      const t = this.tabRun;
      const run = t && t.sheet === sheet && t.at.r === r && t.at.c === c ? t : null;
      this.tabRun = null;
      if (dc) {
        this.move(0, dc);
        this.tabRun = { sheet, col: run?.col ?? c, at: { ...this.doc.sel.active } };
        return;
      }
      if (dr > 0 && run && run.col !== c) {
        const m = this.doc.sheet.mergeAt(r, c);
        this.selectCell(Math.min(MAX_ROWS - 1, (m ? m.r2 : r) + 1), run.col);
        return;
      }
      this.move(dr, dc);
      return;
    }
    let { r, c } = sel.active;
    if (dr) {
      r += dr;
      if (r > rg.r2) {
        r = rg.r1;
        c = c + 1 > rg.c2 ? rg.c1 : c + 1;
      } else if (r < rg.r1) {
        r = rg.r2;
        c = c - 1 < rg.c1 ? rg.c2 : c - 1;
      }
    } else {
      c += dc;
      if (c > rg.c2) {
        c = rg.c1;
        r = r + 1 > rg.r2 ? rg.r1 : r + 1;
      } else if (c < rg.c1) {
        c = rg.c2;
        r = r - 1 < rg.r1 ? rg.r2 : r - 1;
      }
    }
    this.doc.sel = { ranges: sel.ranges, active: { r, c }, anchor: sel.anchor };
    this.ensureVisible(r, c);
    this.emit();
  }

  /** Name box / Go To: accepts A1, A1:B5, Sheet!A1, defined names. */
  goTo(text: string): boolean {
    const doc = this.doc;
    let t = text.trim();
    if (!t) return false;
    let sheetIdx = doc.wb.activeSheet;
    const bang = t.lastIndexOf('!');
    if (bang > 0) {
      const name = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
      const idx = doc.wb.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
      if (idx < 0) return false;
      sheetIdx = idx;
      t = t.slice(bang + 1);
    }
    let rg = parseRange(t);
    if (!rg) {
      const dn = doc.wb.names.find((n) => n.name.toLowerCase() === t.toLowerCase());
      if (dn) return this.goTo(dn.ref);
      return false;
    }
    doc.setActiveSheet(sheetIdx);
    rg = this.withMerges(rg);
    this.select(rg, { r: rg.r1, c: rg.c1 });
    this.ensureVisible(rg.r1, rg.c1);
    return true;
  }

  /* ---------------------------------------------------------- editing */

  startEdit(initial: string | null, mode: 'enter' | 'edit', source: 'cell' | 'bar' = 'cell'): boolean {
    const doc = this.doc;
    const sheet = doc.sheet;
    const { r, c } = doc.sel.active;
    if (this.isLocked(r, c)) {
      this.onError?.('This cell is protected.', 'Unprotect the sheet (Review ▸ Unprotect sheet) to change it.');
      return false;
    }
    const text = initial ?? doc.editText(sheet, r, c);
    this.edit = { sheet: sheet.id, r, c, text, caret: text.length, mode, source, original: doc.editText(sheet, r, c) };
    this.ensureVisible(r, c);
    this.emit();
    return true;
  }

  isLocked(r: number, c: number): boolean {
    const sheet = this.doc.sheet;
    if (!sheet.protection?.enabled) return false;
    return this.doc.styleOf(sheet, r, c).locked !== false;
  }

  updateEdit(text: string, caret: number, source?: 'cell' | 'bar'): void {
    if (!this.edit) return;
    const e = this.edit;
    // typing anywhere but inside the pointed reference ends point mode
    const point = e.point && caret >= e.point.start && caret <= e.point.end && text.length === e.text.length ? e.point : undefined;
    this.edit = { ...e, text, caret, source: source ?? e.source, point };
    this.emit();
  }

  cancelEdit(): void {
    if (!this.edit) return;
    this.edit = null;
    this.emit();
  }

  /** Commits the edit; returns false (and keeps editing) if the formula is invalid. */
  commitEdit(opts: { all?: boolean } = {}): boolean {
    const e = this.edit;
    if (!e) return true;
    const doc = this.doc;
    const idx = doc.wb.sheets.findIndex((s) => s.id === e.sheet);
    if (idx >= 0 && idx !== doc.wb.activeSheet) doc.setActiveSheet(idx);
    if (e.text === e.original && !opts.all) {
      this.edit = null;
      this.emit();
      return true;
    }
    const v = this.validate(e.text, e.r, e.c);
    if (v) {
      this.onError?.(v.title, v.message);
      if (v.block) return false;
    }
    try {
      doc.sel = { ...doc.sel, active: { r: e.r, c: e.c } };
      doc.enter(e.text, { all: opts.all });
    } catch (err) {
      if (err instanceof FormulaError) {
        this.onError?.('There’s a problem with this formula', `${err.message}. Formulas start with = and use commas between arguments, e.g. =SUM(A1:A10) or =IF(A1>5,"Yes","No").`);
        return false;
      }
      throw err;
    }
    this.edit = null;
    this.emit();
    return true;
  }

  /** Data validation check for typed input. */
  validate(text: string, r: number, c: number): { title: string; message: string; block: boolean } | null {
    const doc = this.doc;
    const sheet = doc.sheet;
    const rule = sheet.validations.find((v) => v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
    if (!rule || rule.type === 'any' || text.startsWith('=')) return null;
    if (text === '' && rule.allowBlank !== false) return null;
    const { cell } = doc.interpret(text, sheet);
    const v = cell.v;
    let ok = true;
    const nums = (rule.values ?? []).map((x) => {
      if (x.startsWith('=')) {
        const res = doc.engine.evaluateText(x, sheet.id, r, c);
        return typeof res === 'number' ? res : Number(res);
      }
      return Number(doc.interpret(x, sheet).cell.v);
    });
    const cmp = (n: number) => {
      const [a, b] = nums;
      switch (rule.op ?? 'between') {
        case 'between':
          return n >= Math.min(a, b) && n <= Math.max(a, b);
        case 'notBetween':
          return n < Math.min(a, b) || n > Math.max(a, b);
        case 'equal':
          return n === a;
        case 'notEqual':
          return n !== a;
        case 'greater':
          return n > a;
        case 'less':
          return n < a;
        case 'greaterEqual':
          return n >= a;
        case 'lessEqual':
          return n <= a;
      }
      return true;
    };
    switch (rule.type) {
      case 'list': {
        const items = listItems(doc, rule.values, r, c);
        ok = items.some((it) => it.toLowerCase() === text.trim().toLowerCase());
        break;
      }
      case 'whole':
        ok = typeof v === 'number' && Number.isInteger(v) && cmp(v);
        break;
      case 'decimal':
      case 'date':
      case 'time':
        ok = typeof v === 'number' && cmp(v);
        break;
      case 'textLength':
        ok = cmp(text.length);
        break;
      case 'checkbox':
        ok = typeof v === 'boolean';
        break;
      case 'custom': {
        const res = doc.engine.evaluateText(rule.values[0] ?? 'TRUE', sheet.id, r, c, { r: rule.ranges[0].r1, c: rule.ranges[0].c1 });
        ok = res === true || (typeof res === 'number' && res !== 0);
        break;
      }
    }
    if (ok) return null;
    return {
      title: rule.errorTitle || 'This value doesn’t match the data validation rules',
      message: rule.error || describeRule(rule.type, rule.op, rule.values),
      block: (rule.errorStyle ?? 'stop') === 'stop',
    };
  }

  /* ------------------------------------------------------- point mode */

  /** Whether the caret sits where a reference can be inserted (after =, (, operators…). */
  canPoint(): boolean {
    const e = this.edit;
    if (!e || !e.text.startsWith('=')) return false;
    if (e.point) return true;
    const before = e.text.slice(0, e.caret).trimEnd();
    return /[=(,+\-*/^&<>:;{]$/.test(before);
  }

  /** Inserts or replaces the pointed reference. */
  pointAt(rg: Range, anchor: CellPos, cursor: CellPos, sheetId?: number): void {
    const e = this.edit;
    if (!e) return;
    const doc = this.doc;
    const targetSheet = sheetId ?? doc.sheet.id;
    let ref = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? addrName(rg.r1, rg.c1) : `${addrName(rg.r1, rg.c1)}:${addrName(rg.r2, rg.c2)}`;
    if (rg.r1 === 0 && rg.r2 === MAX_ROWS - 1) ref = `${colName(rg.c1)}:${colName(rg.c2)}`;
    else if (rg.c1 === 0 && rg.c2 === MAX_COLS - 1) ref = `${rg.r1 + 1}:${rg.r2 + 1}`;
    if (targetSheet !== e.sheet) ref = `${quoteSheet(doc.wb.sheetById(targetSheet)?.name ?? '')}!${ref}`;
    const start = e.point ? e.point.start : e.caret;
    const end = e.point ? e.point.end : e.caret;
    const text = e.text.slice(0, start) + ref + e.text.slice(end);
    this.edit = { ...e, text, caret: start + ref.length, point: { start, end: start + ref.length, anchor, cursor } };
    this.emit();
  }

  /** Arrow keys while pointing move the pointed cell. */
  pointMove(dr: number, dc: number, extend: boolean): void {
    const e = this.edit;
    if (!e) return;
    const g = this.geo();
    const cur = e.point?.cursor ?? { r: e.r, c: e.c };
    const anchor = extend ? e.point?.anchor ?? cur : undefined;
    let r = cur.r;
    let c = cur.c;
    if (dr) r = g.rows.step(r, dr > 0 ? 1 : -1);
    if (dc) c = g.cols.step(c, dc > 0 ? 1 : -1);
    const a = anchor ?? { r, c };
    const rg = { r1: Math.min(a.r, r), c1: Math.min(a.c, c), r2: Math.max(a.r, r), c2: Math.max(a.c, c) };
    this.pointAt(rg, a, { r, c });
    this.ensureVisible(r, c);
  }

  toggleAbsolute(): void {
    const e = this.edit;
    if (!e || !e.text.startsWith('=')) return;
    const res = cycleAbsolute(e.text.slice(1), Math.max(0, e.caret - 1));
    this.edit = { ...e, text: '=' + res.text, caret: res.caret + 1, point: undefined };
    this.emit();
  }

  /** References in the formula being edited, for coloured highlights. */
  editRefs(): { range: Range; color: string; start: number; end: number }[] {
    const e = this.edit;
    if (!e || !e.text.startsWith('=')) return [];
    const out: { range: Range; color: string; start: number; end: number }[] = [];
    const toks = tokenize(e.text.slice(1));
    const sheetName = this.doc.wb.sheetById(e.sheet)?.name ?? '';
    const current = this.doc.sheet.name;
    const sig = toks.filter((t) => t.type !== 'ws');
    const seen = new Map<string, string>();
    for (let i = 0; i < sig.length; i++) {
      const t = sig[i];
      if (t.type !== 'ref' && t.type !== 'colrange' && t.type !== 'rowrange') continue;
      let text = t.refText!;
      let end = t.end;
      if (t.type === 'ref' && sig[i + 1]?.type === 'colon' && sig[i + 2]?.type === 'ref' && !sig[i + 2].sheet) {
        text += ':' + sig[i + 2].refText;
        end = sig[i + 2].end;
        i += 2;
      }
      const onSheet = (t.sheet ?? sheetName).toLowerCase() === current.toLowerCase();
      const rg = parseRange(text);
      if (!rg || !onSheet) continue;
      const key = text.replace(/\$/g, '').toUpperCase();
      let color = seen.get(key);
      if (!color) {
        color = REF_COLORS[seen.size % REF_COLORS.length];
        seen.set(key, color);
      }
      out.push({ range: rg, color, start: t.start + 1, end: end + 1 });
    }
    return out;
  }

  cellAddress(): string {
    const sel = this.doc.sel;
    const rg = sel.ranges[sel.ranges.length - 1];
    if (rg.r1 === 0 && rg.r2 === MAX_ROWS - 1 && rg.c1 === 0 && rg.c2 === MAX_COLS - 1) return 'All';
    if (rg.r1 === 0 && rg.r2 === MAX_ROWS - 1) return `${colName(rg.c1)}:${colName(rg.c2)}`;
    if (rg.c1 === 0 && rg.c2 === MAX_COLS - 1) return `${rg.r1 + 1}:${rg.r2 + 1}`;
    const merged = this.doc.sheet.mergeAt(sel.active.r, sel.active.c);
    if (merged && merged.r1 === rg.r1 && merged.c1 === rg.c1 && merged.r2 === rg.r2 && merged.c2 === rg.c2) return addrName(rg.r1, rg.c1);
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return addrName(sel.active.r, sel.active.c);
    return `${addrName(rg.r1, rg.c1)}:${addrName(rg.r2, rg.c2)}`;
  }

  key(r: number, c: number): number {
    return cellKey(r, c);
  }
}

/** Items of a list validation (literal list or range reference). */
export function listItems(doc: SheetDoc, values: string[], r: number, c: number): string[] {
  const sheet = doc.sheet;
  if (values.length === 1 && values[0].startsWith('=')) {
    const res = doc.engine.evaluateText(values[0], sheet.id, r, c);
    if (res && typeof res === 'object' && 'kind' in res && res.kind === 'matrix') {
      return res.data.flat().filter((v) => v !== null && v !== '').map((v) => String(v));
    }
    return res === null || res === undefined ? [] : [String(res)];
  }
  if (values.length === 1 && values[0].includes(',')) return values[0].split(',').map((s) => s.trim()).filter(Boolean);
  return values;
}

function describeRule(type: string, op: string | undefined, values: string[]): string {
  const what: Record<string, string> = { whole: 'a whole number', decimal: 'a number', date: 'a date', time: 'a time', textLength: 'text whose length is', list: 'one of the items in the list', custom: 'a value allowed by the rule' };
  const ops: Record<string, string> = {
    between: `between ${values[0]} and ${values[1]}`,
    notBetween: `not between ${values[0]} and ${values[1]}`,
    equal: `equal to ${values[0]}`,
    notEqual: `not equal to ${values[0]}`,
    greater: `greater than ${values[0]}`,
    less: `less than ${values[0]}`,
    greaterEqual: `at least ${values[0]}`,
    lessEqual: `at most ${values[0]}`,
  };
  if (type === 'list' || type === 'custom') return `Please enter ${what[type]}.`;
  return `Please enter ${what[type] ?? 'a valid value'} ${ops[op ?? 'between'] ?? ''}.`;
}

