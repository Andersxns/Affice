import { adjustFormulaMove, translateFormula } from '../engine/refshift';
import type { Scalar } from '../engine/values';
import { isDateFormat, parseInput, partsToSerial, serialToParts } from '../format/numfmt';
import { cellKey, colName, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { cloneCell } from '../model/workbook';
import type { Cell, CellStyle } from '../model/types';
import type { ClipData, SheetDoc } from '../doc';

/* ================================================================ copy */

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function styleCss(st: CellStyle): string {
  const css: string[] = [];
  if (st.bold) css.push('font-weight:bold');
  if (st.italic) css.push('font-style:italic');
  const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) css.push(`text-decoration:${deco}`);
  if (st.color) css.push(`color:${st.color}`);
  if (st.fill) css.push(`background:${st.fill}`);
  if (st.font) css.push(`font-family:'${st.font}'`);
  if (st.size) css.push(`font-size:${st.size}pt`);
  if (st.hAlign && st.hAlign !== 'general') css.push(`text-align:${st.hAlign === 'centerContinuous' ? 'center' : st.hAlign}`);
  if (st.vAlign) css.push(`vertical-align:${st.vAlign}`);
  const side = (b: CellStyle['bt']) => (b ? `${b.style === 'medium' ? 2 : b.style === 'thick' ? 3 : 1}px ${b.style === 'dashed' ? 'dashed' : b.style === 'dotted' ? 'dotted' : b.style === 'double' ? 'double' : 'solid'} ${b.color ?? '#000'}` : '');
  if (st.bt) css.push(`border-top:${side(st.bt)}`);
  if (st.bb) css.push(`border-bottom:${side(st.bb)}`);
  if (st.bl) css.push(`border-left:${side(st.bl)}`);
  if (st.br) css.push(`border-right:${side(st.br)}`);
  if (st.wrap) css.push('white-space:normal');
  return css.join(';');
}

/** Copies (or cuts) the selection into the internal clipboard and returns text/html for the system clipboard. */
export function copySelection(doc: SheetDoc, cut = false): { text: string; html: string; range: Range } | null {
  const sheet = doc.sheet;
  const ranges = doc.sel.ranges;
  const rg0 = ranges[ranges.length - 1];
  const rg = doc.clampToUsed(sheet, rg0);
  const rows = rg.r2 - rg.r1 + 1;
  const cols = rg.c2 - rg.c1 + 1;
  if (rows * cols > 2_000_000) return null;
  const cells: (Cell | undefined)[][] = [];
  const lines: string[] = [];
  const htmlRows: string[] = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    if (sheet.isRowHidden(r) && !cut) {
      // Excel skips filtered rows when copying
      if (sheet.filterHidden.has(r)) continue;
    }
    const row: (Cell | undefined)[] = [];
    const texts: string[] = [];
    const tds: string[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) {
      const cell = sheet.get(r, c);
      row.push(cloneCell(cell));
      const d = doc.displayText(sheet, r, c);
      const t = d.text;
      texts.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
      const st = doc.styleOf(sheet, r, c);
      const m = sheet.mergeAt(r, c);
      if (m && (m.r1 !== r || m.c1 !== c)) continue;
      const span = m ? ` rowspan="${m.r2 - m.r1 + 1}" colspan="${m.c2 - m.c1 + 1}"` : '';
      const v = d.value;
      const numAttr = typeof v === 'number' ? ` data-num="${v}"` : '';
      tds.push(`<td${span}${numAttr} style="${styleCss(st)}">${escapeHtml(t).replace(/\n/g, '<br>')}</td>`);
    }
    cells.push(row);
    lines.push(texts.join('\t'));
    htmlRows.push(`<tr style="height:${sheet.rowHeight(r)}px">${tds.join('')}</tr>`);
  }
  const colWidths = Array.from({ length: cols }, (_, i) => sheet.colWidth(rg.c1 + i));
  const text = lines.join('\r\n') + (lines.length > 1 ? '\r\n' : '');
  const colgroup = colWidths.map((w) => `<col style="width:${w}px">`).join('');
  const html = `<html><head><meta charset="utf-8"></head><body><table style="border-collapse:collapse;font-family:Calibri,Carlito,sans-serif;font-size:11pt" data-affice="1"><colgroup>${colgroup}</colgroup>${htmlRows.join('')}</table></body></html>`;
  const merges = sheet.merges.filter((m) => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2).map((m) => ({ r1: m.r1 - rg.r1, c1: m.c1 - rg.c1, r2: m.r2 - rg.r1, c2: m.c2 - rg.c1 }));
  const clip: ClipData = { sheet: sheet.id, range: rg, cells, cut, text, colWidths, merges };
  doc.clip = clip;
  doc.emit('view');
  return { text, html, range: rg };
}

/* =============================================================== paste */

export type PasteMode = 'all' | 'values' | 'formulas' | 'formats' | 'transpose' | 'valuesAndFormats' | 'noBorders' | 'colWidths' | 'link';

interface Block {
  cells: (Cell | undefined)[][];
  merges: Range[];
  /** Formula translation base (internal copies). */
  from?: { sheet: number; r: number; c: number };
  colWidths?: number[];
}

function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let q = false;
  const t = text.replace(/\r\n?/g, '\n');
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = false;
      } else cur += ch;
      continue;
    }
    if (ch === '"' && cur === '') q = true;
    else if (ch === '\t') {
      row.push(cur);
      cur = '';
    } else if (ch === '\n') {
      row.push(cur);
      rows.push(row);
      row = [];
      cur = '';
    } else cur += ch;
  }
  if (cur !== '' || row.length) {
    row.push(cur);
    rows.push(row);
  }
  return rows;
}

function cssColor(v: string | null | undefined): string | undefined {
  if (!v) return undefined;
  const s = v.trim();
  if (!s || s === 'transparent' || s === 'inherit' || s === 'initial') return undefined;
  const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/.exec(s);
  if (m) {
    if (m[4] !== undefined && Number(m[4]) === 0) return undefined;
    return '#' + [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('');
  }
  if (/^#[0-9a-f]{3,8}$/i.test(s)) return s.toLowerCase();
  if (/^[a-z]+$/i.test(s)) return s.toLowerCase();
  return undefined;
}

/** Reads an HTML table (Excel, Google Sheets, web pages, Affice) into cells + styles. */
function parseHtmlTable(html: string, doc: SheetDoc): Block | null {
  const dom = new DOMParser().parseFromString(html, 'text/html');
  const table = dom.querySelector('table');
  if (!table) return null;
  const grid: (Cell | undefined)[][] = [];
  const merges: Range[] = [];
  const taken = new Set<string>();
  const trs = [...table.querySelectorAll('tr')];
  trs.forEach((tr, ri) => {
    let ci = 0;
    grid[ri] = grid[ri] ?? [];
    for (const td of [...tr.children].filter((x) => x.tagName === 'TD' || x.tagName === 'TH')) {
      while (taken.has(`${ri}:${ci}`)) ci++;
      const el = td as HTMLElement;
      const rs = Math.max(1, Number(el.getAttribute('rowspan') ?? 1));
      const cs = Math.max(1, Number(el.getAttribute('colspan') ?? 1));
      for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) taken.add(`${ri + a}:${ci + b}`);
      if (rs > 1 || cs > 1) merges.push({ r1: ri, c1: ci, r2: ri + rs - 1, c2: ci + cs - 1 });
      const text = (el.innerText ?? el.textContent ?? '').replace(/ /g, ' ').replace(/\n$/, '');
      const numAttr = el.getAttribute('data-num') ?? el.getAttribute('x:num') ?? el.getAttribute('sdval');
      const style: CellStyle = {};
      const st = el.style;
      const fw = st.fontWeight || (el.querySelector('b,strong') ? 'bold' : '');
      if (fw === 'bold' || Number(fw) >= 600 || el.tagName === 'TH') style.bold = true;
      if (st.fontStyle === 'italic' || el.querySelector('i,em')) style.italic = true;
      if (/underline/.test(st.textDecoration) || el.querySelector('u')) style.underline = 'single';
      if (/line-through/.test(st.textDecoration) || el.querySelector('s,strike,del')) style.strike = true;
      const color = cssColor(st.color);
      if (color && color !== '#000000' && color !== 'black') style.color = color;
      const bg = cssColor(st.backgroundColor || st.background) ?? cssColor(el.getAttribute('bgcolor'));
      if (bg && bg !== '#ffffff' && bg !== 'white') style.fill = bg;
      const ta = st.textAlign || el.getAttribute('align');
      if (ta === 'center' || ta === 'right' || ta === 'left' || ta === 'justify') style.hAlign = ta;
      if (/normal|pre-wrap/.test(st.whiteSpace) && /\n/.test(text)) style.wrap = true;
      const size = parseFloat(st.fontSize);
      if (size && /pt$/.test(st.fontSize) && Math.abs(size - 11) > 0.1) style.size = size;
      const numFmt = el.getAttribute('x:str') !== null ? '@' : undefined;
      let cell: Cell | undefined;
      if (numAttr !== null && numAttr !== '' && !Number.isNaN(Number(numAttr))) cell = { v: Number(numAttr) };
      else if (text !== '') {
        const p = numFmt === '@' ? { value: text } : parseInput(text);
        cell = { v: p.value === null ? undefined : p.value };
        if ('format' in p && p.format) style.numFmt = p.format;
      }
      const sid = Object.keys(style).length ? doc.wb.styleId(style) : 0;
      if (sid) cell = { ...(cell ?? {}), s: sid };
      grid[ri][ci] = cell;
      ci += cs;
    }
  });
  const width = Math.max(0, ...grid.map((r) => r.length));
  for (const r of grid) while (r.length < width) r.push(undefined);
  return { cells: grid, merges };
}

/**
 * Pastes at the active cell. Internal copies keep formulas (adjusted) and formatting;
 * external data comes from HTML tables or tab-separated text.
 */
export function paste(doc: SheetDoc, sys: { text?: string; html?: string }, mode: PasteMode = 'all'): Range | null {
  const sheet = doc.sheet;
  const clip = doc.clip;
  let block: Block | null = null;
  const sysText = (sys.text ?? '').replace(/\r\n/g, '\n').replace(/\n$/, '');
  const internal = clip && (sys.text === undefined || sysText === clip.text.replace(/\r\n/g, '\n').replace(/\n$/, ''));
  if (internal && clip) {
    block = { cells: clip.cells, merges: clip.merges, from: { sheet: clip.sheet, r: clip.range.r1, c: clip.range.c1 }, colWidths: clip.colWidths };
  } else if (sys.html && /<table/i.test(sys.html)) block = parseHtmlTable(sys.html, doc);
  if (!block && sys.text !== undefined) {
    const rows = parseTsv(sys.text);
    block = {
      cells: rows.map((row) =>
        row.map((t) => {
          if (t === '') return undefined;
          if (t.startsWith('=') && t.length > 1) {
            try {
              return { f: doc.interpret(t, sheet).cell.f };
            } catch {
              return { v: t };
            }
          }
          const p = parseInput(t);
          const cell: Cell = { v: p.value === null ? undefined : p.value };
          if (p.format) cell.s = doc.wb.styleId({ numFmt: p.format });
          return cell;
        }),
      ),
      merges: [],
    };
  }
  if (!block || !block.cells.length) return null;
  if (mode === 'transpose') block = transposeBlock(block);
  const bh = block.cells.length;
  const bw = Math.max(1, ...block.cells.map((r) => r.length));
  const { active } = doc.sel;
  const selRg = doc.sel.ranges[doc.sel.ranges.length - 1];
  // tile the block over a larger selection that is an exact multiple of it (like Excel)
  const selH = selRg.r2 - selRg.r1 + 1;
  const selW = selRg.c2 - selRg.c1 + 1;
  const tileR = selH > bh && selH % bh === 0 && selH <= 100000 ? selH / bh : 1;
  const tileC = selW > bw && selW % bw === 0 && selW <= 1000 ? selW / bw : 1;
  const r0 = tileR > 1 || tileC > 1 ? selRg.r1 : active.r;
  const c0 = tileR > 1 || tileC > 1 ? selRg.c1 : active.c;
  const dest: Range = { r1: r0, c1: c0, r2: Math.min(MAX_ROWS - 1, r0 + bh * tileR - 1), c2: Math.min(MAX_COLS - 1, c0 + bw * tileC - 1) };
  const isCut = internal && clip?.cut;
  const srcSheet = isCut ? doc.wb.sheetById(clip!.sheet) : undefined;

  doc.transact(isCut ? 'Move' : 'Paste', (tx) => {
    if (isCut && srcSheet && clip) {
      // clear the source first (the destination may overlap it)
      for (let r = clip.range.r1; r <= clip.range.r2; r++) for (let c = clip.range.c1; c <= clip.range.c2; c++) tx.set(srcSheet, cellKey(r, c), undefined);
      if (srcSheet.merges.some((m) => m.r1 >= clip.range.r1 && m.r2 <= clip.range.r2 && m.c1 >= clip.range.c1 && m.c2 <= clip.range.c2)) {
        tx.prop(srcSheet, 'merges');
        srcSheet.merges = srcSheet.merges.filter((m) => !(m.r1 >= clip.range.r1 && m.r2 <= clip.range.r2 && m.c1 >= clip.range.c1 && m.c2 <= clip.range.c2));
      }
    }
    for (let tr = 0; tr < tileR; tr++)
      for (let tc = 0; tc < tileC; tc++)
        for (let i = 0; i < bh; i++)
          for (let j = 0; j < bw; j++) {
            const r = r0 + tr * bh + i;
            const c = c0 + tc * bw + j;
            if (r >= MAX_ROWS || c >= MAX_COLS) continue;
            const src = block!.cells[i]?.[j];
            const key = cellKey(r, c);
            const cur = sheet.cells.get(key);
            const next = pasteCell(doc, src, cur, mode, block!.from ? { dr: r - (block!.from.r + (mode === 'transpose' ? j : i)), dc: c - (block!.from.c + (mode === 'transpose' ? i : j)), cut: !!isCut } : null, block!.from ? { sheet: block!.from.sheet, r: block!.from.r + (mode === 'transpose' ? j : i), c: block!.from.c + (mode === 'transpose' ? i : j) } : null);
            if (next === 'keep') continue;
            tx.set(sheet, key, next);
          }
    if ((mode === 'all' || mode === 'formats' || mode === 'transpose') && block!.merges.length) {
      tx.prop(sheet, 'merges');
      sheet.merges = sheet.merges.filter((m) => !(m.r1 <= dest.r2 && m.r2 >= dest.r1 && m.c1 <= dest.c2 && m.c2 >= dest.c1));
      for (let tr = 0; tr < tileR; tr++)
        for (let tc = 0; tc < tileC; tc++)
          for (const m of block!.merges) sheet.merges.push({ r1: r0 + tr * bh + m.r1, c1: c0 + tc * bw + m.c1, r2: r0 + tr * bh + m.r2, c2: c0 + tc * bw + m.c2 });
    }
    if (mode === 'colWidths' && block!.colWidths) {
      tx.prop(sheet, 'cols');
      block!.colWidths.forEach((w, j) => sheet.cols.set(c0 + j, { ...(sheet.cols.get(c0 + j) ?? {}), w }));
    }
    if (isCut && srcSheet && clip) {
      // references to the moved block follow it
      const op = { srcSheet: srcSheet.name, src: clip.range, dstSheet: sheet.name, dr: r0 - clip.range.r1, dc: c0 - clip.range.c1 };
      for (const s of doc.wb.sheets)
        for (const [k, cell] of s.cells) {
          if (cell.f === undefined) continue;
          const nf = adjustFormulaMove(cell.f, s.name, op);
          if (nf !== cell.f) tx.patch(s, k, { f: nf });
        }
      tx.rebuild = true;
    }
    doc.sel = { ranges: [dest], active: { r: r0, c: c0 }, anchor: { r: r0, c: c0 } };
  });
  if (isCut) doc.clip = null;
  return dest;
}

function transposeBlock(b: Block): Block {
  const h = b.cells.length;
  const w = Math.max(0, ...b.cells.map((r) => r.length));
  const cells: (Cell | undefined)[][] = [];
  for (let j = 0; j < w; j++) {
    const row: (Cell | undefined)[] = [];
    for (let i = 0; i < h; i++) row.push(b.cells[i]?.[j]);
    cells.push(row);
  }
  return { cells, merges: b.merges.map((m) => ({ r1: m.c1, c1: m.r1, r2: m.c2, c2: m.r2 })), from: b.from, colWidths: undefined };
}

function pasteCell(
  doc: SheetDoc,
  src: Cell | undefined,
  cur: Cell | undefined,
  mode: PasteMode,
  shift: { dr: number; dc: number; cut: boolean } | null,
  origin: { sheet: number; r: number; c: number } | null,
): Cell | undefined | 'keep' {
  const wb = doc.wb;
  const formula = (f: string) => (shift && !shift.cut ? translateFormula(f, shift.dr, shift.dc) : f);
  const computed = (): Scalar | undefined => {
    if (!src) return undefined;
    if (src.f !== undefined && origin) {
      const s = wb.sheetById(origin.sheet);
      return s ? doc.engine.displayValue(s.id, origin.r, origin.c) ?? undefined : src.v;
    }
    return src.v;
  };
  switch (mode) {
    case 'all':
    case 'transpose':
      return src ? { ...cloneCell(src), f: src.f !== undefined ? formula(src.f) : undefined, v: src.f !== undefined ? undefined : src.v } : undefined;
    case 'noBorders': {
      if (!src) return undefined;
      const st = { ...wb.style(src.s) };
      delete st.bt;
      delete st.bb;
      delete st.bl;
      delete st.br;
      return { ...cloneCell(src), f: src.f !== undefined ? formula(src.f) : undefined, s: wb.styleId(st) || undefined };
    }
    case 'values': {
      const v = computed();
      if (v === undefined && !cur) return 'keep';
      return { ...(cur ?? {}), v: v ?? undefined, f: undefined };
    }
    case 'valuesAndFormats':
      return src ? { v: computed(), s: src.s, link: src.link } : cur?.s ? undefined : 'keep';
    case 'formulas':
      if (!src) return cur ? { ...cur, v: undefined, f: undefined } : 'keep';
      return { ...(cur ?? {}), f: src.f !== undefined ? formula(src.f) : undefined, v: src.f !== undefined ? undefined : src.v };
    case 'formats':
      return { ...(cur ?? {}), s: src?.s };
    case 'link': {
      if (!origin) return 'keep';
      const s = wb.sheetById(origin.sheet);
      const ref = `${s && s !== doc.sheet ? `'${s.name.replace(/'/g, "''")}'!` : ''}${colName(origin.c)}${origin.r + 1}`;
      return { ...(cur ?? {}), f: ref, v: undefined };
    }
    case 'colWidths':
      return 'keep';
  }
}

/* ================================================================ move */

/** Drag-and-drop move of a block (references follow the cells). */
export function moveBlock(doc: SheetDoc, src: Range, dr: number, dc: number, copy = false): void {
  if (!dr && !dc) return;
  const sheet = doc.sheet;
  const dest: Range = { r1: src.r1 + dr, c1: src.c1 + dc, r2: src.r2 + dr, c2: src.c2 + dc };
  if (dest.r1 < 0 || dest.c1 < 0 || dest.r2 >= MAX_ROWS || dest.c2 >= MAX_COLS) return;
  const data: { r: number; c: number; cell: Cell | undefined }[] = [];
  for (let r = src.r1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) data.push({ r, c, cell: cloneCell(sheet.get(r, c)) });
  doc.transact(copy ? 'Copy cells' : 'Move cells', (tx) => {
    if (!copy) for (const d of data) tx.set(sheet, cellKey(d.r, d.c), undefined);
    for (const d of data) {
      let cell = d.cell;
      if (cell?.f !== undefined && copy) cell = { ...cell, f: translateFormula(cell.f, dr, dc), v: undefined };
      tx.set(sheet, cellKey(d.r + dr, d.c + dc), cell);
    }
    const inner = sheet.merges.filter((m) => m.r1 >= src.r1 && m.r2 <= src.r2 && m.c1 >= src.c1 && m.c2 <= src.c2);
    if (inner.length) {
      tx.prop(sheet, 'merges');
      const moved = inner.map((m) => ({ r1: m.r1 + dr, c1: m.c1 + dc, r2: m.r2 + dr, c2: m.c2 + dc }));
      sheet.merges = [...(copy ? sheet.merges : sheet.merges.filter((m) => !inner.includes(m))).filter((m) => !(m.r1 <= dest.r2 && m.r2 >= dest.r1 && m.c1 <= dest.c2 && m.c2 >= dest.c1)), ...moved];
    }
    if (!copy) {
      const op = { srcSheet: sheet.name, src, dstSheet: sheet.name, dr, dc };
      for (const s of doc.wb.sheets)
        for (const [k, cell] of s.cells) {
          if (cell.f === undefined) continue;
          const nf = adjustFormulaMove(cell.f, s.name, op);
          if (nf !== cell.f) tx.patch(s, k, { f: nf });
        }
    }
    doc.sel = { ranges: [dest], active: { r: dest.r1, c: dest.c1 }, anchor: { r: dest.r1, c: dest.c1 } };
  });
}

/* ============================================================ autofill */

const LISTS: string[][] = [
  ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
  ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['Q1', 'Q2', 'Q3', 'Q4'],
];

function matchCase(template: string, word: string): string {
  if (template === template.toUpperCase()) return word.toUpperCase();
  if (template === template.toLowerCase()) return word.toLowerCase();
  return word;
}

type Gen = (k: number) => Cell | undefined; // k = 0,1,2… positions after the source

/** Builds a generator continuing a line of source cells. */
function seriesFor(doc: SheetDoc, src: (Cell | undefined)[], dir: { dr: number; dc: number }): Gen {
  const n = src.length;
  const values = src.map((c) => (c?.f === undefined ? c?.v : undefined));
  const allNum = n > 0 && src.every((c) => c && c.f === undefined && typeof c.v === 'number');
  const fmt = doc.wb.style(src[0]?.s).numFmt;
  const cyclic: Gen = (k) => {
    const i = k % n;
    const cell = src[i];
    if (!cell) return undefined;
    const steps = Math.floor(k / n) + 1;
    if (cell.f !== undefined) return { ...cloneCell(cell), f: translateFormula(cell.f, dir.dr * n * steps, dir.dc * n * steps), v: undefined };
    return cloneCell(cell);
  };
  if (allNum) {
    const xs = values as number[];
    if (n === 1) {
      if (isDateFormat(fmt)) return (k) => ({ ...cloneCell(src[0]), v: xs[0] + (k + 1) });
      return cyclic;
    }
    // monthly/yearly date steps keep the day of month
    if (isDateFormat(fmt)) {
      const parts = xs.map((x) => serialToParts(x));
      const monthStep = (parts[1].y - parts[0].y) * 12 + (parts[1].m - parts[0].m);
      const sameDay = parts.every((p) => p.d === parts[0].d) && monthStep !== 0;
      const consistent = parts.every((p, i) => i === 0 || (p.y - parts[i - 1].y) * 12 + (p.m - parts[i - 1].m) === monthStep);
      if (sameDay && consistent) {
        const last = parts[n - 1];
        return (k) => {
          const total = last.y * 12 + (last.m - 1) + monthStep * (k + 1);
          const y = Math.floor(total / 12);
          const m = (total % 12) + 1;
          const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
          return { ...cloneCell(src[(k + n) % n]), v: partsToSerial(y, m, Math.min(last.d, dim)) };
        };
      }
    }
    // least-squares linear trend (like Excel's AutoFill)
    const mx = (n - 1) / 2;
    const my = xs.reduce((a, b) => a + b, 0) / n;
    let sxy = 0;
    let sxx = 0;
    xs.forEach((y, i) => {
      sxy += (i - mx) * (y - my);
      sxx += (i - mx) * (i - mx);
    });
    const step = sxx ? sxy / sxx : 0;
    const icpt = my - step * mx;
    return (k) => ({ ...cloneCell(src[k % n]), v: parseFloat((icpt + step * (n + k)).toPrecision(15)) });
  }
  // lists (weekdays, months, quarters)
  if (values.every((v) => typeof v === 'string' && v !== '')) {
    const strs = values as string[];
    for (const list of LISTS) {
      const idx = strs.map((s) => list.findIndex((w) => w.toLowerCase() === s.toLowerCase()));
      if (idx.every((i) => i >= 0)) {
        const step = n > 1 ? (((idx[1] - idx[0]) % list.length) + list.length) % list.length || 1 : 1;
        return (k) => ({ ...cloneCell(src[k % n]), v: matchCase(strs[0], list[(idx[n - 1] + step * (k + 1)) % list.length]) });
      }
    }
    // text with a trailing (or leading) number: "Item 1", "Q3 2024" → increments the last number
    const m = strs.map((s) => /^(.*?)(\d+)(\D*)$/.exec(s));
    if (m.every(Boolean) && m.every((x) => x![1] === m[0]![1] && x![3] === m[0]![3])) {
      const nums = m.map((x) => Number(x![2]));
      const step = n > 1 ? nums[1] - nums[0] || 1 : 1;
      const width = m[0]![2].length;
      return (k) => {
        const v = nums[n - 1] + step * (k + 1);
        const digits = m[0]![2].startsWith('0') ? String(Math.max(0, v)).padStart(width, '0') : String(v);
        return { ...cloneCell(src[k % n]), v: `${m[0]![1]}${digits}${m[0]![3]}` };
      };
    }
  }
  return cyclic;
}

/**
 * Fill-handle drag: extends `src` to cover `target` (which contains it).
 * Dragging back inside the selection clears the uncovered cells.
 */
export function autoFill(doc: SheetDoc, src: Range, target: Range, copyOnly = false): void {
  const sheet = doc.sheet;
  doc.transact('Fill', (tx) => {
    if (target.r2 < src.r2 || target.c2 < src.c2) {
      // shrink: clear what the drag uncovered
      for (let r = src.r1; r <= src.r2; r++)
        for (let c = src.c1; c <= src.c2; c++) {
          if (r <= target.r2 && c <= target.c2) continue;
          const k = cellKey(r, c);
          if (sheet.cells.get(k)?.v !== undefined || sheet.cells.get(k)?.f !== undefined) tx.patch(sheet, k, { v: undefined, f: undefined });
        }
      doc.sel = { ranges: [target], active: { r: target.r1, c: target.c1 }, anchor: { r: target.r1, c: target.c1 } };
      return;
    }
    const vertical = target.r1 !== src.r1 || target.r2 !== src.r2;
    const forward = vertical ? target.r2 > src.r2 : target.c2 > src.c2;
    const lines = vertical ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1;
    const len = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
    const extra = vertical ? (forward ? target.r2 - src.r2 : src.r1 - target.r1) : forward ? target.c2 - src.c2 : src.c1 - target.c1;
    for (let l = 0; l < lines; l++) {
      const cells: (Cell | undefined)[] = [];
      for (let i = 0; i < len; i++) {
        const idx = forward ? i : len - 1 - i;
        const r = vertical ? src.r1 + idx : src.r1 + l;
        const c = vertical ? src.c1 + l : src.c1 + idx;
        cells.push(cloneCell(sheet.get(r, c)));
      }
      const dir = vertical ? { dr: forward ? 1 : -1, dc: 0 } : { dr: 0, dc: forward ? 1 : -1 };
      const gen = copyOnly
        ? (k: number) => {
            const cell = cells[k % len];
            if (cell?.f !== undefined) {
              const steps = Math.floor(k / len) + 1;
              return { ...cell, f: translateFormula(cell.f, dir.dr * len * steps, dir.dc * len * steps), v: undefined };
            }
            return cloneCell(cell);
          }
        : seriesFor(doc, cells, dir);
      for (let k = 0; k < extra; k++) {
        const step = len + k;
        const r = vertical ? (forward ? src.r1 + step : src.r2 - step) : src.r1 + l;
        const c = vertical ? src.c1 + l : forward ? src.c1 + step : src.c2 - step;
        if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) continue;
        tx.set(sheet, cellKey(r, c), gen(k));
      }
    }
    doc.sel = { ranges: [target], active: doc.sel.active, anchor: doc.sel.anchor };
  });
}

/** Ctrl+D / Ctrl+R: copy the first row/column of the selection into the rest. */
export function fillDirection(doc: SheetDoc, dir: 'down' | 'right'): void {
  const sheet = doc.sheet;
  doc.transact(dir === 'down' ? 'Fill down' : 'Fill right', (tx) => {
    for (const rg0 of doc.sel.ranges) {
      let rg = doc.clampToUsed(sheet, rg0, 1);
      if (dir === 'down' && rg.r1 === rg.r2 && rg.r1 > 0) rg = { ...rg, r1: rg.r1 - 1 };
      if (dir === 'right' && rg.c1 === rg.c2 && rg.c1 > 0) rg = { ...rg, c1: rg.c1 - 1 };
      for (let r = rg.r1; r <= rg.r2; r++)
        for (let c = rg.c1; c <= rg.c2; c++) {
          const sr = dir === 'down' ? rg.r1 : r;
          const sc = dir === 'right' ? rg.c1 : c;
          if (r === sr && c === sc) continue;
          const cell = cloneCell(sheet.get(sr, sc));
          if (cell?.f !== undefined) cell.f = translateFormula(cell.f, r - sr, c - sc);
          if (cell?.f !== undefined) delete cell.v;
          tx.set(sheet, cellKey(r, c), cell);
        }
    }
  });
}

/* ============================================================ flash fill */

/**
 * Flash Fill (Ctrl+E): learns a text transformation from the examples typed
 * in the active column and fills the remaining rows next to the data.
 */
export function flashFill(doc: SheetDoc): number {
  const sheet = doc.sheet;
  const { c } = doc.sel.active;
  if (c === 0) return 0;
  const ext = doc.engine.extent(sheet.id);
  const srcCols: number[] = [];
  for (let cc = c - 1; cc >= 0 && sheet.get(doc.sel.active.r, cc) !== undefined; cc--) srcCols.unshift(cc);
  if (!srcCols.length) srcCols.push(c - 1);
  const rowText = (r: number) => srcCols.map((cc) => doc.displayText(sheet, r, cc).text).join(' ');
  // find the data block rows
  let top = doc.sel.active.r;
  while (top > 0 && rowText(top - 1).trim() !== '') top--;
  let bottom = doc.sel.active.r;
  while (bottom < ext.rows && rowText(bottom + 1).trim() !== '') bottom++;
  const examples: { input: string; output: string }[] = [];
  for (let r = top; r <= bottom; r++) {
    const out = sheet.get(r, c)?.v;
    if (typeof out === 'string' && out !== '') examples.push({ input: rowText(r), output: out });
  }
  if (!examples.length) return 0;
  const program = learnTransform(examples);
  if (!program) return 0;
  let filled = 0;
  doc.transact('Flash Fill', (tx) => {
    for (let r = top; r <= bottom; r++) {
      const cur = sheet.get(r, c);
      if (cur?.v !== undefined || cur?.f !== undefined) continue;
      const out = program(rowText(r));
      if (out === null) continue;
      tx.patch(sheet, cellKey(r, c), { v: out });
      filled++;
    }
  });
  return filled;
}

type Program = (input: string) => string | null;

/** Tries a family of token-based programs and keeps the first consistent with every example. */
function learnTransform(examples: { input: string; output: string }[]): Program | null {
  const splitWords = (s: string) => s.split(/[\s,;]+/).filter(Boolean);
  const cases: ((w: string) => string)[] = [(w) => w, (w) => w.toUpperCase(), (w) => w.toLowerCase(), (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()];
  const candidates: Program[] = [];
  const ex = examples[0];
  const words = splitWords(ex.input);
  const outWords = ex.output.split(/(\s+|[.,@\-_/]+)/).filter((x) => x !== '');
  // map each output token to (word index, case, prefix-length) or literal
  type Piece = { lit: string } | { word: number; fromEnd: boolean; kase: number; initial: boolean };
  const options: Piece[][] = outWords.map((tok) => {
    const opts: Piece[] = [];
    words.forEach((w, i) => {
      cases.forEach((f, k) => {
        if (f(w) === tok) {
          opts.push({ word: i, fromEnd: false, kase: k, initial: false });
          opts.push({ word: words.length - 1 - i, fromEnd: true, kase: k, initial: false });
        }
        if (tok.length === 1 && f(w.charAt(0)) === tok) {
          opts.push({ word: i, fromEnd: false, kase: k, initial: true });
          opts.push({ word: words.length - 1 - i, fromEnd: true, kase: k, initial: true });
        }
      });
    });
    opts.push({ lit: tok });
    return opts;
  });
  // email-ish / whole-cell extraction candidates
  candidates.push((s) => s.match(/[\w.+-]+@[\w-]+\.[\w.]+/)?.[0] ?? null);
  candidates.push((s) => s.match(/\d+(?:[.,]\d+)?/)?.[0] ?? null);
  // enumerate combinations (bounded)
  const combos: Piece[][] = [];
  const walk = (i: number, acc: Piece[]) => {
    if (combos.length > 400) return;
    if (i === options.length) {
      combos.push(acc);
      return;
    }
    for (const o of options[i].slice(0, 12)) walk(i + 1, [...acc, o]);
  };
  walk(0, []);
  for (const combo of combos) {
    candidates.push((s) => {
      const ws = splitWords(s);
      let out = '';
      for (const p of combo) {
        if ('lit' in p) out += p.lit;
        else {
          const w = ws[p.fromEnd ? ws.length - 1 - p.word : p.word];
          if (w === undefined) return null;
          out += cases[p.kase](p.initial ? w.charAt(0) : w);
        }
      }
      return out;
    });
  }
  return candidates.find((prog) => examples.every((e) => prog(e.input) === e.output)) ?? null;
}
