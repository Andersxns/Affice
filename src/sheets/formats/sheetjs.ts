/* eslint-disable @typescript-eslint/no-explicit-any */
import { FErr, isErr, asErr, type Scalar } from '../engine/values';
import { SHORT_DATE } from '../format/numfmt';
import { cellKey, keyCol, keyRow, parseRange } from '../model/address';
import { Workbook } from '../model/workbook';
import type { CellStyle } from '../model/types';
import type { SheetDoc } from '../doc';
import { fromExcelFormula, toExcelFormula } from './excelFormula';

let mod: Promise<typeof import('xlsx')> | null = null;
const sheetjs = () => (mod ??= import('xlsx'));

const ERROR_CODES: Record<number, string> = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA' };

function colorOf(c: any): string | undefined {
  const rgb = c?.rgb;
  if (typeof rgb !== 'string' || rgb.length < 6) return undefined;
  return '#' + rgb.slice(-6).toLowerCase();
}

/** Reads .xls, .xlsb, .numbers and HTML tables through SheetJS. */
export async function importWithSheetJS(bytes: Uint8Array): Promise<{ wb: Workbook; warning?: string }> {
  const X = await sheetjs();
  const xwb = X.read(bytes, { type: 'array', cellFormula: true, cellStyles: true, cellNF: true, cellDates: false, dense: false });
  const wb = new Workbook();
  xwb.SheetNames.forEach((name: string, i: number) => {
    const ws: any = xwb.Sheets[name];
    const sheet = wb.addSheet(wb.uniqueSheetName(name));
    sheet.hidden = !!(xwb as any).Workbook?.Sheets?.[i]?.Hidden;
    sheet.defaultColWidth = 64;
    sheet.defaultRowHeight = 20;
    for (const key of Object.keys(ws)) {
      if (key[0] === '!') continue;
      const addr = X.utils.decode_cell(key);
      const c = ws[key];
      const out: { v?: Scalar; f?: string; s?: number; link?: string; note?: string } = {};
      switch (c.t) {
        case 'n':
          out.v = c.v;
          break;
        case 's':
        case 'str':
          out.v = String(c.v ?? '');
          break;
        case 'b':
          out.v = !!c.v;
          break;
        case 'e':
          out.v = new FErr((ERROR_CODES[c.v] ?? c.w ?? '#VALUE!') as FErr['error']);
          break;
        case 'd':
          out.v = c.v instanceof Date ? c.v.getTime() / 86400000 + 25569 : Number(c.v);
          break;
        default:
          break;
      }
      if (c.f) out.f = fromExcelFormula(String(c.f).replace(/^=/, ''));
      const st: CellStyle = {};
      if (c.z && typeof c.z === 'string' && !/^general$/i.test(c.z)) st.numFmt = c.z === 'm/d/yy' ? SHORT_DATE : c.z;
      const s = c.s;
      if (s) {
        const fill = colorOf(s.fgColor);
        if (fill && s.patternType && s.patternType !== 'none') st.fill = fill;
        if (s.font) {
          if (s.font.bold) st.bold = true;
          if (s.font.italic) st.italic = true;
          if (s.font.underline) st.underline = 'single';
          const fc = colorOf(s.font.color);
          if (fc && fc !== '#000000') st.color = fc;
          if (s.font.name && s.font.name !== 'Calibri' && s.font.name !== 'Arial') st.font = s.font.name;
          if (s.font.sz && Number(s.font.sz) !== 11 && Number(s.font.sz) !== 10) st.size = Number(s.font.sz);
        }
        if (s.alignment?.horizontal) st.hAlign = s.alignment.horizontal;
        if (s.alignment?.wrapText) st.wrap = true;
      }
      if (Object.keys(st).length) out.s = wb.styleId(st);
      if (c.l?.Target) out.link = c.l.Target;
      if (Array.isArray(c.c) && c.c.length) out.note = c.c.map((x: any) => x.t).join('\n');
      if (out.v === undefined) delete out.v;
      if (out.v !== undefined || out.f !== undefined || out.s || out.link || out.note) sheet.put(cellKey(addr.r, addr.c), out);
    }
    for (const m of ws['!merges'] ?? []) sheet.merges.push({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c });
    (ws['!cols'] ?? []).forEach((col: any, c: number) => {
      if (!col) return;
      const w = col.wpx ?? (col.wch !== undefined ? col.wch * 7 + 5 : col.width !== undefined ? col.width * 7 + 5 : undefined);
      const info: { w?: number; hidden?: boolean } = {};
      if (w !== undefined && Math.abs(w - 64) > 1) info.w = Math.round(w);
      if (col.hidden) info.hidden = true;
      if (Object.keys(info).length) sheet.cols.set(c, info);
    });
    (ws['!rows'] ?? []).forEach((row: any, r: number) => {
      if (!row) return;
      const h = row.hpx ?? (row.hpt !== undefined ? (row.hpt * 96) / 72 : undefined);
      const info: { h?: number; hidden?: boolean; custom?: boolean } = {};
      if (h !== undefined && Math.abs(h - 20) > 1) {
        info.h = Math.round(h);
        info.custom = true;
      }
      if (row.hidden) info.hidden = true;
      if (Object.keys(info).length) sheet.rows.set(r, info);
    });
    if (ws['!autofilter']?.ref) {
      const rg = parseRange(ws['!autofilter'].ref);
      if (rg) sheet.filter = { range: rg, columns: {} };
    }
  });
  for (const n of (xwb as any).Workbook?.Names ?? []) {
    if (!n.Name || n.Name.startsWith('_xlnm')) continue;
    wb.names.push({ name: n.Name, ref: fromExcelFormula(String(n.Ref ?? '')), sheet: typeof n.Sheet === 'number' ? wb.sheets[n.Sheet]?.id : undefined });
  }
  if (!wb.sheets.length) wb.addSheet();
  return { wb };
}

/** Writes .xls (BIFF8) through SheetJS: values, formulas, number formats, merges and sizes. */
export async function exportWithSheetJS(doc: SheetDoc, bookType: 'xls' | 'ods' | 'fods' | 'xlsb'): Promise<Uint8Array> {
  const X = await sheetjs();
  const xwb = X.utils.book_new();
  for (const sheet of doc.wb.sheets) {
    const ws: any = {};
    let maxR = 0;
    let maxC = 0;
    for (const [k, cell] of sheet.cells) {
      const r = keyRow(k);
      const c = keyCol(k);
      const v = doc.value(sheet, r, c);
      if (v === undefined && cell.f === undefined) continue;
      const out: any = {};
      if (isErr(v)) {
        out.t = 'e';
        out.v = Number(Object.entries(ERROR_CODES).find(([, e]) => e === asErr(v).error)?.[0] ?? 15);
        out.w = asErr(v).error;
      } else if (typeof v === 'number') {
        out.t = 'n';
        out.v = v;
      } else if (typeof v === 'boolean') {
        out.t = 'b';
        out.v = v;
      } else {
        out.t = 's';
        out.v = v ?? '';
      }
      if (cell.f !== undefined) out.f = toExcelFormula(cell.f);
      const fmt = doc.wb.style(cell.s).numFmt;
      if (fmt) out.z = fmt;
      if (cell.link) out.l = { Target: cell.link };
      if (cell.note) out.c = [{ a: 'Affice', t: cell.note }];
      ws[X.utils.encode_cell({ r, c })] = out;
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    }
    ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
    if (sheet.merges.length) ws['!merges'] = sheet.merges.map((m) => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
    const cols: any[] = [];
    for (let c = 0; c <= maxC; c++) {
      const info = sheet.cols.get(c);
      cols.push({ wpx: info?.w ?? sheet.defaultColWidth, hidden: info?.hidden || undefined });
    }
    ws['!cols'] = cols;
    const rows: any[] = [];
    for (const [r, info] of sheet.rows) rows[r] = { hpx: info.h, hidden: info.hidden || undefined };
    ws['!rows'] = rows;
    X.utils.book_append_sheet(xwb, ws, sheet.name.slice(0, 31));
  }
  const out = X.write(xwb, { bookType, type: 'array' }) as ArrayBuffer;
  return new Uint8Array(out);
}
