import { cellKey, keyCol, keyRow } from '../model/address';
import type { Sheet } from '../model/workbook';
import type { SheetDoc } from '../doc';
import { cellFont, fitText, measure } from './render';

let scratch: CanvasRenderingContext2D | null = null;
function ctx2d(): CanvasRenderingContext2D {
  if (!scratch) scratch = document.createElement('canvas').getContext('2d')!;
  return scratch;
}

/** Best width (px, zoom 1) for a column based on its contents. */
export function autoFitWidth(doc: SheetDoc, sheet: Sheet, c: number, rows?: [number, number]): number {
  const ctx = ctx2d();
  let best = 0;
  let n = 0;
  for (const [k] of sheet.cells) {
    if (keyCol(k) !== c) continue;
    const r = keyRow(k);
    if (rows && (r < rows[0] || r > rows[1])) continue;
    if (sheet.mergeAt(r, c)) continue;
    const st = doc.styleOf(sheet, r, c);
    if (st.wrap) continue;
    const d = doc.displayText(sheet, r, c);
    if (!d.text) continue;
    const { font } = cellFont(st, 1);
    const text = typeof d.value === 'number' ? fitText(ctx, font, d.value, st.numFmt, 1e9, d.text) : d.text;
    const w = Math.max(...text.split('\n').map((line) => measure(ctx, font, line))) + 8 + (st.indent ?? 0) * 9;
    if (w > best) best = w;
    if (++n > 20000) break;
  }
  if (sheet.filter && sheet.filter.range.c1 <= c && sheet.filter.range.c2 >= c) best += 18;
  return best ? Math.min(1200, Math.ceil(best)) : sheet.defaultColWidth;
}

/** Height needed by a row's contents (px, zoom 1), or null when the default fits. */
export function autoFitHeight(doc: SheetDoc, sheet: Sheet, r: number): number | null {
  const ctx = ctx2d();
  let best = 0;
  for (let c = 0; c <= doc.engine.extent(sheet.id).cols; c++) {
    const k = cellKey(r, c);
    const cell = sheet.cells.get(k);
    const st = doc.styleOf(sheet, r, c);
    if (!cell && !st.size) continue;
    const { font, px } = cellFont(st, 1);
    let h = px * 1.22 + 6;
    if (st.wrap && cell) {
      const d = doc.displayText(sheet, r, c);
      if (d.text) {
        const width = sheet.colWidth(c) - 6;
        let lines = 0;
        for (const para of d.text.split('\n')) {
          let line = '';
          lines++;
          for (const w of para.split(/(\s+)/)) {
            const cand = line + w;
            if (line && measure(ctx, font, cand.trimEnd()) > width) {
              lines++;
              line = w.trimStart();
            } else line = cand;
          }
        }
        h = lines * px * 1.22 + 6;
      }
    }
    if ((st.rotation ?? 0) !== 0 && cell) {
      const d = doc.displayText(sheet, r, c);
      const w = measure(ctx, font, d.text);
      h = Math.max(h, Math.abs(Math.sin(((st.rotation === 255 ? 90 : st.rotation!) * Math.PI) / 180)) * w + 8);
    }
    if (h > best) best = h;
  }
  const need = Math.ceil(best);
  return need > sheet.defaultRowHeight ? Math.min(409, need) : null;
}

/** Re-fits rows that don't have a user-set height (after edits or font changes). */
export function autoFitRows(doc: SheetDoc, sheet: Sheet, rows: Iterable<number>): boolean {
  let changed = false;
  let n = 0;
  for (const r of rows) {
    if (++n > 2000) break;
    const info = sheet.rows.get(r);
    if (info?.custom || info?.hidden) continue;
    const h = autoFitHeight(doc, sheet, r);
    if ((info?.h ?? null) === h) continue;
    if (h === null) {
      if (info) {
        const next = { ...info };
        delete next.h;
        if (Object.keys(next).length) sheet.rows.set(r, next);
        else sheet.rows.delete(r);
      }
    } else sheet.rows.set(r, { ...(info ?? {}), h });
    changed = true;
  }
  return changed;
}
