import { adjustFormulaStructural, deleteSheetInFormula, renameSheetInFormula, shiftRange, type StructuralOp } from '../engine/refshift';
import { cellKey, keyCol, keyRow, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { cloneCell, Sheet } from '../model/workbook';
import type { Cell, ColInfo, RowInfo } from '../model/types';
import type { SheetDoc, TxBuilder } from '../doc';

/* ======================================================= rows/columns */

function shiftIndex(i: number, at: number, count: number): number | null {
  if (count > 0) return i >= at ? i + count : i;
  const n = -count;
  if (i >= at && i < at + n) return null;
  return i >= at + n ? i - n : i;
}

/** Rewrites every formula, name and range in the workbook for an insert/delete. */
function adjustWorkbook(doc: SheetDoc, tx: TxBuilder, sheet: Sheet, op: StructuralOp): void {
  const wb = doc.wb;
  for (const s of wb.sheets) {
    for (const [k, cell] of s.cells) {
      if (cell.f === undefined) continue;
      const nf = adjustFormulaStructural(cell.f, s.name, op);
      if (nf === cell.f) continue;
      if (s === sheet) cell.f = nf; // whole sheet already snapshotted
      else tx.patch(s, k, { f: nf });
    }
    // chart sources and CF/validation formulas
    if (s.charts.some((ch) => adjustFormulaStructural(ch.source, s.name, op) !== ch.source)) {
      tx.prop(s, 'charts');
      for (const ch of s.charts) ch.source = adjustFormulaStructural(ch.source, s.name, op);
    }
  }
  if (wb.names.some((n) => adjustFormulaStructural(n.ref, n.sheet !== undefined ? wb.sheetById(n.sheet)?.name ?? sheet.name : sheet.name, op) !== n.ref)) {
    tx.names();
    for (const n of wb.names) n.ref = adjustFormulaStructural(n.ref, n.sheet !== undefined ? wb.sheetById(n.sheet)?.name ?? sheet.name : sheet.name, op);
  }
}

function shiftSheetRanges(tx: TxBuilder, sheet: Sheet, axis: 'row' | 'col', at: number, count: number): void {
  tx.prop(sheet, 'merges');
  sheet.merges = sheet.merges.map((m) => shiftRange(m, axis, at, count)).filter((m): m is Range => !!m && (m.r1 !== m.r2 || m.c1 !== m.c2));
  if (sheet.cf.length) {
    tx.prop(sheet, 'cf');
    sheet.cf = sheet.cf
      .map((cf) => ({ ...cf, ranges: cf.ranges.map((r) => shiftRange(r, axis, at, count)).filter((r): r is Range => !!r) }))
      .filter((cf) => cf.ranges.length);
  }
  if (sheet.validations.length) {
    tx.prop(sheet, 'validations');
    sheet.validations = sheet.validations
      .map((v) => ({ ...v, ranges: v.ranges.map((r) => shiftRange(r, axis, at, count)).filter((r): r is Range => !!r) }))
      .filter((v) => v.ranges.length);
  }
  const moveAnchor = (a: { r: number; c: number }) => {
    const i = axis === 'row' ? a.r : a.c;
    const n = count > 0 ? (i >= at ? i + count : i) : i >= at - count ? i + count : i >= at ? at : i;
    return axis === 'row' ? { ...a, r: Math.max(0, n) } : { ...a, c: Math.max(0, n) };
  };
  if (sheet.charts.length) {
    tx.prop(sheet, 'charts');
    for (const ch of sheet.charts) ch.anchor = { ...ch.anchor, ...moveAnchor(ch.anchor) };
  }
  if (sheet.images.length) {
    tx.prop(sheet, 'images');
    for (const im of sheet.images) im.anchor = { ...im.anchor, ...moveAnchor(im.anchor) };
  }
  if (sheet.filter) {
    tx.prop(sheet, 'filter');
    const r = shiftRange(sheet.filter.range, axis, at, count);
    if (!r) sheet.filter = undefined;
    else {
      const cols: typeof sheet.filter.columns = {};
      if (axis === 'col') {
        for (const [k, v] of Object.entries(sheet.filter.columns)) {
          const abs = sheet.filter.range.c1 + Number(k);
          const n = shiftIndex(abs, at, count);
          if (n !== null) cols[n - r.c1] = v;
        }
      }
      sheet.filter = { ...sheet.filter, range: r, columns: axis === 'col' ? cols : sheet.filter.columns };
    }
  }
}

export function insertRows(doc: SheetDoc, at: number, count: number): void {
  structural(doc, 'row', at, count, `Insert ${count === 1 ? 'row' : 'rows'}`);
}

export function deleteRows(doc: SheetDoc, at: number, count: number): void {
  structural(doc, 'row', at, -count, `Delete ${count === 1 ? 'row' : 'rows'}`);
}

export function insertCols(doc: SheetDoc, at: number, count: number): void {
  structural(doc, 'col', at, count, `Insert ${count === 1 ? 'column' : 'columns'}`);
}

export function deleteCols(doc: SheetDoc, at: number, count: number): void {
  structural(doc, 'col', at, -count, `Delete ${count === 1 ? 'column' : 'columns'}`);
}

function structural(doc: SheetDoc, axis: 'row' | 'col', at: number, count: number, label: string): void {
  const sheet = doc.sheet;
  doc.transact(label, (tx) => {
    tx.prop(sheet, 'cells');
    tx.prop(sheet, axis === 'row' ? 'rows' : 'cols');
    tx.rebuild = true;
    const moved = new Map<number, Cell>();
    for (const [k, cell] of sheet.cells) {
      let r = keyRow(k);
      let c = keyCol(k);
      if (axis === 'row') {
        const n = shiftIndex(r, at, count);
        if (n === null || n >= MAX_ROWS) continue;
        r = n;
      } else {
        const n = shiftIndex(c, at, count);
        if (n === null || n >= MAX_COLS) continue;
        c = n;
      }
      moved.set(cellKey(r, c), cell);
    }
    // inserted rows/columns inherit the formatting of the one before (like Excel)
    if (count > 0 && at > 0) {
      for (const [k, cell] of sheet.cells) {
        const r = keyRow(k);
        const c = keyCol(k);
        if ((axis === 'row' ? r : c) !== at - 1 || !cell.s) continue;
        for (let i = 0; i < count; i++) {
          const nk = axis === 'row' ? cellKey(at + i, c) : cellKey(r, at + i);
          if (!moved.has(nk)) moved.set(nk, { s: cell.s });
        }
      }
    }
    sheet.cells = moved;
    sheet.markExtentStale();
    const infos = (axis === 'row' ? sheet.rows : sheet.cols) as Map<number, RowInfo & ColInfo>;
    const next = new Map<number, RowInfo & ColInfo>();
    for (const [i, info] of infos) {
      const n = shiftIndex(i, at, count);
      if (n !== null && n < (axis === 'row' ? MAX_ROWS : MAX_COLS)) next.set(n, info);
    }
    if (count > 0 && at > 0) {
      const prev = infos.get(at - 1);
      if (prev && (prev.s || prev.h || prev.w)) for (let i = 0; i < count; i++) next.set(at + i, { ...prev, hidden: false });
    }
    if (axis === 'row') sheet.rows = next;
    else sheet.cols = next;
    shiftSheetRanges(tx, sheet, axis, at, count);
    adjustWorkbook(doc, tx, sheet, { sheet: sheet.name, axis, at, count });
    // keep the selection sensible
    const sel = doc.sel;
    if (count < 0) {
      const a = sel.active;
      const pos = axis === 'row' ? { r: Math.min(a.r, MAX_ROWS - 1), c: a.c } : { r: a.r, c: Math.min(a.c, MAX_COLS - 1) };
      doc.sel = { ...sel, active: pos, anchor: pos };
    }
  });
}

/* =========================================================== insert/delete cells with shift */

export function insertCellsShift(doc: SheetDoc, rg: Range, dir: 'down' | 'right'): void {
  const sheet = doc.sheet;
  doc.transact('Insert cells', (tx) => {
    tx.prop(sheet, 'cells');
    tx.rebuild = true;
    const moved = new Map<number, Cell>();
    const h = rg.r2 - rg.r1 + 1;
    const w = rg.c2 - rg.c1 + 1;
    for (const [k, cell] of sheet.cells) {
      let r = keyRow(k);
      let c = keyCol(k);
      if (dir === 'down' && c >= rg.c1 && c <= rg.c2 && r >= rg.r1) r += h;
      if (dir === 'right' && r >= rg.r1 && r <= rg.r2 && c >= rg.c1) c += w;
      if (r < MAX_ROWS && c < MAX_COLS) moved.set(cellKey(r, c), cell);
    }
    sheet.cells = moved;
    sheet.markExtentStale();
    shiftFormulasInBlock(doc, tx, sheet, rg, dir === 'down' ? h : 0, dir === 'right' ? w : 0);
  });
}

export function deleteCellsShift(doc: SheetDoc, rg: Range, dir: 'up' | 'left'): void {
  const sheet = doc.sheet;
  doc.transact('Delete cells', (tx) => {
    tx.prop(sheet, 'cells');
    tx.rebuild = true;
    const moved = new Map<number, Cell>();
    const h = rg.r2 - rg.r1 + 1;
    const w = rg.c2 - rg.c1 + 1;
    for (const [k, cell] of sheet.cells) {
      let r = keyRow(k);
      let c = keyCol(k);
      if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) continue;
      if (dir === 'up' && c >= rg.c1 && c <= rg.c2 && r > rg.r2) r -= h;
      if (dir === 'left' && r >= rg.r1 && r <= rg.r2 && c > rg.c2) c -= w;
      moved.set(cellKey(r, c), cell);
    }
    sheet.cells = moved;
    sheet.markExtentStale();
    shiftFormulasInBlock(doc, tx, sheet, rg, dir === 'up' ? -h : 0, dir === 'left' ? -w : 0);
  });
}

/** Approximates Excel's reference update for partial shifts: references to moved cells follow them. */
function shiftFormulasInBlock(doc: SheetDoc, tx: TxBuilder, sheet: Sheet, rg: Range, dr: number, dc: number): void {
  const op: StructuralOp = dr ? { sheet: sheet.name, axis: 'row', at: dr > 0 ? rg.r1 : rg.r1, count: dr } : { sheet: sheet.name, axis: 'col', at: rg.c1, count: dc };
  // Only references inside the shifted band (same columns for vertical shifts) are affected.
  for (const s of doc.wb.sheets)
    for (const [k, cell] of s.cells) {
      if (cell.f === undefined) continue;
      const nf = adjustFormulaStructural(cell.f, s.name, op);
      if (nf === cell.f) continue;
      if (s === sheet) cell.f = nf;
      else tx.patch(s, k, { f: nf });
    }
}

/* ============================================================ sizing */

export function setColWidth(doc: SheetDoc, cols: number[], w: number): void {
  const sheet = doc.sheet;
  doc.transact('Column width', (tx) => {
    tx.prop(sheet, 'cols');
    for (const c of cols) sheet.cols.set(c, { ...(sheet.cols.get(c) ?? {}), w: Math.max(0, Math.round(w)), hidden: w <= 0 ? true : undefined });
  });
}

export function setRowHeight(doc: SheetDoc, rows: number[], h: number, custom = true): void {
  const sheet = doc.sheet;
  doc.transact('Row height', (tx) => {
    tx.prop(sheet, 'rows');
    for (const r of rows) sheet.rows.set(r, { ...(sheet.rows.get(r) ?? {}), h: Math.max(0, Math.round(h)), custom: custom || undefined, hidden: h <= 0 ? true : undefined });
  });
}

export function setHidden(doc: SheetDoc, axis: 'row' | 'col', idx: number[], hidden: boolean): void {
  const sheet = doc.sheet;
  doc.transact(hidden ? 'Hide' : 'Unhide', (tx) => {
    tx.prop(sheet, axis === 'row' ? 'rows' : 'cols');
    const map = axis === 'row' ? sheet.rows : sheet.cols;
    for (const i of idx) {
      const info = map.get(i) ?? {};
      const next = { ...info, hidden: hidden || undefined };
      if (!hidden && (next as RowInfo).h === 0) delete (next as RowInfo).h;
      if (!hidden && (next as ColInfo).w === 0) delete (next as ColInfo).w;
      map.set(i, next);
    }
  });
}

/** Rows/columns covered by the selection (full span of each range). */
export function selectedIndexes(doc: SheetDoc, axis: 'row' | 'col'): number[] {
  const out = new Set<number>();
  const ext = doc.engine.extent(doc.sheet.id);
  for (const rg of doc.sel.ranges) {
    const [a, b] = axis === 'row' ? [rg.r1, rg.r2] : [rg.c1, rg.c2];
    const limit = axis === 'row' ? Math.max(ext.rows + 100, a + 1000) : Math.max(ext.cols + 50, a + 200);
    for (let i = a; i <= Math.min(b, limit); i++) out.add(i);
  }
  return [...out].sort((x, y) => x - y);
}

export function setFreeze(doc: SheetDoc, rows: number, cols: number): void {
  const sheet = doc.sheet;
  doc.transact('Freeze panes', (tx) => {
    tx.prop(sheet, 'freeze');
    sheet.view.freezeRows = Math.max(0, rows);
    sheet.view.freezeCols = Math.max(0, cols);
  });
}

/* ============================================================= sheets */

const BAD_NAME = /[\\/?*[\]:]/;

export function validateSheetName(doc: SheetDoc, name: string, except?: Sheet): string | null {
  const n = name.trim();
  if (!n) return 'The name can’t be empty.';
  if (n.length > 31) return 'Sheet names can have at most 31 characters.';
  if (BAD_NAME.test(n)) return 'Sheet names can’t contain \\ / ? * [ ] :';
  if (n.startsWith("'") || n.endsWith("'")) return 'Sheet names can’t start or end with an apostrophe.';
  const clash = doc.wb.sheetByName(n);
  if (clash && clash !== except) return 'Another sheet already has that name.';
  return null;
}

export function addSheet(doc: SheetDoc, index?: number, name?: string): Sheet {
  let created!: Sheet;
  doc.transact('Insert sheet', (tx) => {
    tx.sheets();
    created = doc.wb.addSheet(name ? doc.wb.uniqueSheetName(name) : undefined, index ?? doc.wb.activeSheet + 1);
    doc.wb.activeSheet = doc.wb.sheets.indexOf(created);
  });
  return created;
}

export function deleteSheet(doc: SheetDoc, index: number): boolean {
  const wb = doc.wb;
  if (wb.sheets.filter((s) => !s.hidden).length <= 1 && !wb.sheets[index]?.hidden) return false;
  const victim = wb.sheets[index];
  if (!victim) return false;
  doc.transact('Delete sheet', (tx) => {
    tx.sheets();
    wb.sheets.splice(index, 1);
    for (const s of wb.sheets)
      for (const [k, cell] of s.cells) {
        if (cell.f === undefined) continue;
        const nf = deleteSheetInFormula(cell.f, victim.name);
        if (nf !== cell.f) tx.patch(s, k, { f: nf });
      }
    if (wb.names.some((n) => n.sheet === victim.id || deleteSheetInFormula(n.ref, victim.name) !== n.ref)) {
      tx.names();
      wb.names = wb.names.filter((n) => n.sheet !== victim.id).map((n) => ({ ...n, ref: deleteSheetInFormula(n.ref, victim.name) }));
    }
    let next = Math.min(index, wb.sheets.length - 1);
    while (wb.sheets[next]?.hidden && next > 0) next--;
    wb.activeSheet = next;
  });
  return true;
}

export function renameSheet(doc: SheetDoc, index: number, name: string): string | null {
  const wb = doc.wb;
  const sheet = wb.sheets[index];
  if (!sheet) return 'No such sheet';
  const err = validateSheetName(doc, name, sheet);
  if (err) return err;
  const old = sheet.name;
  const nn = name.trim();
  if (old === nn) return null;
  doc.transact('Rename sheet', (tx) => {
    tx.prop(sheet, 'name');
    sheet.name = nn;
    for (const s of wb.sheets) {
      for (const [k, cell] of s.cells) {
        if (cell.f === undefined) continue;
        const nf = renameSheetInFormula(cell.f, old, nn);
        if (nf !== cell.f) tx.patch(s, k, { f: nf });
      }
      if (s.charts.some((c) => renameSheetInFormula(c.source, old, nn) !== c.source)) {
        tx.prop(s, 'charts');
        for (const c of s.charts) c.source = renameSheetInFormula(c.source, old, nn);
      }
    }
    if (wb.names.some((n) => renameSheetInFormula(n.ref, old, nn) !== n.ref)) {
      tx.names();
      for (const n of wb.names) n.ref = renameSheetInFormula(n.ref, old, nn);
    }
    tx.rebuild = true;
  });
  return null;
}

export function moveSheet(doc: SheetDoc, from: number, to: number): void {
  const wb = doc.wb;
  if (from === to || from < 0 || to < 0 || from >= wb.sheets.length || to >= wb.sheets.length) return;
  doc.transact('Move sheet', (tx) => {
    tx.sheets();
    const active = wb.sheets[wb.activeSheet];
    const [s] = wb.sheets.splice(from, 1);
    wb.sheets.splice(to, 0, s);
    wb.activeSheet = wb.sheets.indexOf(active);
  });
}

export function duplicateSheet(doc: SheetDoc, index: number): Sheet | null {
  const wb = doc.wb;
  const src = wb.sheets[index];
  if (!src) return null;
  let copy!: Sheet;
  doc.transact('Duplicate sheet', (tx) => {
    tx.sheets();
    copy = wb.addSheet(wb.uniqueSheetName(src.name), index + 1);
    for (const [k, cell] of src.cells) copy.put(k, cloneCell(cell));
    copy.cols = new Map([...src.cols].map(([k, v]) => [k, { ...v }]));
    copy.rows = new Map([...src.rows].map(([k, v]) => [k, { ...v }]));
    copy.merges = structuredClone(src.merges);
    copy.cf = structuredClone(src.cf).map((c) => ({ ...c, id: c.id + '-copy' }));
    copy.validations = structuredClone(src.validations);
    copy.charts = structuredClone(src.charts).map((c) => ({ ...c, id: c.id + '-' + copy.id }));
    copy.images = structuredClone(src.images).map((c) => ({ ...c, id: c.id + '-' + copy.id }));
    copy.filter = src.filter ? structuredClone(src.filter) : undefined;
    copy.view = { ...structuredClone(src.view) };
    copy.tabColor = src.tabColor;
    copy.defaultColWidth = src.defaultColWidth;
    copy.defaultRowHeight = src.defaultRowHeight;
    wb.activeSheet = index + 1;
  });
  return copy;
}

export function setSheetProp(doc: SheetDoc, index: number, prop: 'tabColor' | 'hidden', value: string | boolean | undefined): void {
  const wb = doc.wb;
  const sheet = wb.sheets[index];
  if (!sheet) return;
  if (prop === 'hidden' && value && wb.sheets.filter((s) => !s.hidden).length <= 1) return;
  doc.transact(prop === 'hidden' ? (value ? 'Hide sheet' : 'Unhide sheet') : 'Tab colour', (tx) => {
    tx.prop(sheet, prop);
    (sheet as unknown as Record<string, unknown>)[prop] = value;
    if (prop === 'hidden' && value && wb.activeSheet === index) {
      const next = wb.sheets.findIndex((s) => !s.hidden);
      wb.activeSheet = next < 0 ? 0 : next;
    }
  });
}

