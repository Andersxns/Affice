import type { PdfJob } from '@/shared/types';
import { escapeHtml } from '@/lib/utils';
import { colName, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { DEFAULT_PRINT, type CellStyle, type SheetPrint } from '../model/types';
import type { Sheet } from '../model/workbook';
import type { SheetDoc } from '../doc';
import { buildChartConfig } from '../charts/chartConfig';

const PAPER: Record<SheetPrint['paper'], [number, number]> = {
  letter: [8.5, 11],
  a4: [8.27, 11.69],
  legal: [8.5, 14],
  a3: [11.69, 16.54],
  a5: [5.83, 8.27],
  tabloid: [11, 17],
};
const MARGINS: Record<SheetPrint['margins'], { top: number; right: number; bottom: number; left: number }> = {
  normal: { top: 0.75, right: 0.7, bottom: 0.75, left: 0.7 },
  narrow: { top: 0.5, right: 0.25, bottom: 0.5, left: 0.25 },
  wide: { top: 1, right: 1, bottom: 1, left: 1 },
};

function borderCss(b: CellStyle['bt']): string {
  if (!b) return '';
  const w = b.style === 'medium' ? 2 : b.style === 'thick' || b.style === 'double' ? 3 : 1;
  const kind = b.style === 'dashed' || b.style === 'mediumDashed' ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : b.style === 'double' ? 'double' : 'solid';
  return `${w}px ${kind} ${b.color ?? '#000'}`;
}

function cellCss(st: CellStyle, isNum: boolean, gridlines: boolean): string {
  const css: string[] = [];
  if (st.font) css.push(`font-family:'${st.font}'`);
  if (st.size) css.push(`font-size:${st.size}pt`);
  if (st.bold) css.push('font-weight:700');
  if (st.italic) css.push('font-style:italic');
  const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) css.push(`text-decoration:${deco}`);
  if (st.underline === 'double') css.push('text-decoration-style:double');
  if (st.color) css.push(`color:${st.color}`);
  if (st.fill) css.push(`background:${st.fill}`);
  const h = st.hAlign && st.hAlign !== 'general' ? (st.hAlign === 'centerContinuous' || st.hAlign === 'distributed' ? 'center' : st.hAlign === 'fill' ? 'left' : st.hAlign) : isNum ? 'right' : 'left';
  css.push(`text-align:${h}`);
  css.push(`vertical-align:${st.vAlign === 'middle' ? 'middle' : st.vAlign === 'top' ? 'top' : 'bottom'}`);
  css.push(st.wrap ? 'white-space:pre-wrap;word-wrap:break-word' : 'white-space:pre');
  if (st.indent) css.push(`padding-left:${3 + st.indent * 9}px`);
  const g = gridlines ? '1px solid #d4d4d4' : '';
  const t = borderCss(st.bt) || g;
  const b = borderCss(st.bb) || g;
  const l = borderCss(st.bl) || g;
  const r = borderCss(st.br) || g;
  if (t) css.push(`border-top:${t}`);
  if (b) css.push(`border-bottom:${b}`);
  if (l) css.push(`border-left:${l}`);
  if (r) css.push(`border-right:${r}`);
  return css.join(';');
}

/** The area that gets printed: print area, else used cells plus objects. */
export function printRange(doc: SheetDoc, sheet: Sheet, pr: SheetPrint): Range | null {
  if (pr.area) return pr.area;
  const ext = doc.engine.extent(sheet.id);
  let r2 = ext.rows;
  let c2 = ext.cols;
  for (const m of sheet.merges) {
    r2 = Math.max(r2, m.r2);
    c2 = Math.max(c2, m.c2);
  }
  const objEnd = (a: { r: number; c: number; dx: number; dy: number }, w: number, h: number) => {
    let x = a.dx + w;
    let c = a.c;
    while (x > sheet.colWidth(c) && c < MAX_COLS - 1) {
      x -= sheet.colWidth(c);
      c++;
    }
    let y = a.dy + h;
    let r = a.r;
    while (y > sheet.rowHeight(r) && r < MAX_ROWS - 1) {
      y -= sheet.rowHeight(r);
      r++;
    }
    return { r, c };
  };
  for (const o of [...sheet.charts, ...sheet.images]) {
    const e = objEnd(o.anchor, o.w, o.h);
    r2 = Math.max(r2, e.r);
    c2 = Math.max(c2, e.c);
  }
  if (r2 < 0 || c2 < 0) return null;
  return { r1: 0, c1: 0, r2, c2 };
}

async function chartImage(doc: SheetDoc, sheet: Sheet, spec: Sheet['charts'][number]): Promise<string> {
  const { default: Chart } = await import('chart.js/auto');
  const canvas = document.createElement('canvas');
  const scale = 2;
  canvas.width = spec.w * scale;
  canvas.height = spec.h * scale;
  canvas.style.width = `${spec.w}px`;
  canvas.style.height = `${spec.h}px`;
  const holder = document.createElement('div');
  holder.style.cssText = `position:fixed;left:-10000px;top:0;width:${spec.w}px;height:${spec.h}px`;
  holder.appendChild(canvas);
  document.body.appendChild(holder);
  try {
    const cfg = buildChartConfig(doc, spec, sheet, false, false);
    cfg.options = { ...(cfg.options ?? {}), responsive: false, devicePixelRatio: scale, animation: false };
    const ch = new Chart(canvas, cfg);
    const url = ch.toBase64Image('image/png');
    ch.destroy();
    return url;
  } finally {
    holder.remove();
  }
}

export async function sheetHtml(doc: SheetDoc, sheet: Sheet, rg: Range, pr: SheetPrint): Promise<{ html: string; width: number; height: number }> {
  const cols: number[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) if (!sheet.isColHidden(c)) cols.push(c);
  const rows: number[] = [];
  for (let r = rg.r1; r <= rg.r2; r++) if (!sheet.isRowHidden(r)) rows.push(r);
  const colPx = cols.map((c) => sheet.colWidth(c));
  const headW = pr.headings ? 36 : 0;
  const width = colPx.reduce((a, b) => a + b, 0) + headW;
  const merges = sheet.merges.filter((m) => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2);
  const covered = new Set<string>();
  for (const m of merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r}:${c}`);
  let out = `<table class="sheet-print" style="width:${width}px"><colgroup>${pr.headings ? `<col style="width:${headW}px">` : ''}${colPx.map((w) => `<col style="width:${w}px">`).join('')}</colgroup>`;
  const repeat = Math.min(pr.repeatRows ?? 0, rows.length);
  const rowHtml = (r: number) => {
    let tr = `<tr style="height:${sheet.rowHeight(r)}px">`;
    if (pr.headings) tr += `<th class="rh">${r + 1}</th>`;
    for (const c of cols) {
      if (covered.has(`${r}:${c}`)) continue;
      const m = merges.find((x) => x.r1 === r && x.c1 === c);
      const span = m ? ` rowspan="${rows.filter((x) => x >= m.r1 && x <= m.r2).length}" colspan="${cols.filter((x) => x >= m.c1 && x <= m.c2).length}"` : '';
      const d = doc.displayText(sheet, r, c);
      const st = doc.styleOf(sheet, r, c);
      const color = d.color ? `color:${d.color};` : '';
      const cell = sheet.get(r, c);
      const content = escapeHtml(d.text);
      const body = cell?.link ? `<a href="${escapeHtml(cell.link)}">${content}</a>` : content;
      tr += `<td${span} style="${color}${cellCss(st, typeof d.value === 'number', pr.gridlines)}">${body}</td>`;
    }
    return tr + '</tr>';
  };
  if (pr.headings) out += `<thead><tr><th></th>${cols.map((c) => `<th>${colName(c)}</th>`).join('')}</tr>${rows.slice(0, repeat).map(rowHtml).join('')}</thead>`;
  else if (repeat) out += `<thead>${rows.slice(0, repeat).map(rowHtml).join('')}</thead>`;
  out += `<tbody>${rows.slice(repeat).map(rowHtml).join('')}</tbody></table>`;
  const height = rows.reduce((a, r) => a + sheet.rowHeight(r), 0) + (pr.headings ? 20 : 0);
  // pictures and charts positioned over the table
  const xOf = (c: number, dx: number) => headW + cols.filter((x) => x < c).reduce((a, x) => a + sheet.colWidth(x), 0) + dx;
  const yOf = (r: number, dy: number) => (pr.headings ? 20 : 0) + rows.filter((x) => x < r).reduce((a, x) => a + sheet.rowHeight(x), 0) + dy;
  let objs = '';
  for (const img of sheet.images) objs += `<img class="sheet-obj" src="${img.src}" style="left:${xOf(img.anchor.c, img.anchor.dx)}px;top:${yOf(img.anchor.r, img.anchor.dy)}px;width:${img.w}px;height:${img.h}px">`;
  for (const ch of sheet.charts) {
    try {
      const url = await chartImage(doc, sheet, ch);
      objs += `<img class="sheet-obj" src="${url}" style="left:${xOf(ch.anchor.c, ch.anchor.dx)}px;top:${yOf(ch.anchor.r, ch.anchor.dy)}px;width:${ch.w}px;height:${ch.h}px">`;
    } catch {
      /* skip charts that fail to render */
    }
  }
  return { html: `<div class="sheet-print-wrap" style="width:${width}px">${out}${objs}</div>`, width, height };
}

const PRINT_CSS = `
.sheet-page{break-after:page;page-break-after:always}
.sheet-page:last-child{break-after:auto;page-break-after:auto}
.sheet-print-wrap{position:relative;transform-origin:top left}
.sheet-print{border-collapse:collapse;table-layout:fixed;font-family:Calibri,Carlito,'Liberation Sans',sans-serif;font-size:11pt;color:#000}
.sheet-print td{padding:0 3px;overflow:hidden;line-height:1.2}
.sheet-print th{font:600 9pt Inter,system-ui,sans-serif;background:#f2f2f2;border:1px solid #d4d4d4;color:#555}
.sheet-print tr{break-inside:avoid;page-break-inside:avoid}
.sheet-print a{color:#0563c1}
.sheet-obj{position:absolute}
.sheet-title{font:600 10pt Inter,system-ui,sans-serif;color:#555;margin:0 0 6px}
`;

function hfTemplate(text: string | undefined, title: string, sheetName: string): string | undefined {
  if (!text) return undefined;
  const html = escapeHtml(text)
    .replace(/&amp;P/g, '<span class="pageNumber"></span>')
    .replace(/&amp;N/g, '<span class="totalPages"></span>')
    .replace(/&amp;D/g, '<span class="date"></span>')
    .replace(/&amp;F/g, escapeHtml(title))
    .replace(/&amp;A/g, escapeHtml(sheetName));
  return `<div style="font:9px Inter,system-ui,sans-serif;color:#666;width:100%;text-align:center;padding:0 24px">${html}</div>`;
}

/** Builds the print/PDF job for the active sheet, the selection or the whole workbook. */
export async function buildSheetPdfJob(doc: SheetDoc, scope: 'sheet' | 'selection' | 'workbook', title: string): Promise<PdfJob> {
  const active = doc.sheet;
  const pr: SheetPrint = { ...DEFAULT_PRINT, ...(active.print ?? {}) };
  const [pw, ph] = PAPER[pr.paper];
  const [W, H] = pr.orientation === 'landscape' ? [ph, pw] : [pw, ph];
  const m = MARGINS[pr.margins];
  const printableW = (W - m.left - m.right) * 96;
  const printableH = (H - m.top - m.bottom) * 96;
  const sheets = scope === 'workbook' ? doc.wb.sheets.filter((s) => !s.hidden) : [active];
  let html = '';
  for (const sheet of sheets) {
    const spr: SheetPrint = { ...DEFAULT_PRINT, ...(sheet.print ?? {}), ...(sheet === active ? pr : {}) };
    const rg = scope === 'selection' && sheet === active ? doc.clampToUsed(sheet, doc.sel.ranges[doc.sel.ranges.length - 1]) : printRange(doc, sheet, spr);
    if (!rg) continue;
    const part = await sheetHtml(doc, sheet, rg, spr);
    let scale = spr.fit === 'none' ? spr.scale / 100 : Math.min(1, printableW / part.width);
    if (spr.fit === 'page') scale = Math.min(scale, printableH / Math.max(1, part.height));
    const zoom = Math.max(0.1, scale);
    html += `<section class="sheet-page">${sheets.length > 1 ? `<p class="sheet-title">${escapeHtml(sheet.name)}</p>` : ''}<div style="zoom:${zoom.toFixed(4)};${spr.centerH ? 'margin:0 auto;' : ''}">${part.html}</div></section>`;
  }
  if (!html) html = '<p style="font:12pt sans-serif;color:#888">This sheet is empty.</p>';
  return {
    html,
    css: PRINT_CSS,
    bodyClass: 'print-sheet',
    pageWidthIn: W,
    pageHeightIn: H,
    marginsIn: m,
    headerHtml: hfTemplate(pr.header, title, active.name),
    footerHtml: hfTemplate(pr.footer, title, active.name),
    title,
  };
}
