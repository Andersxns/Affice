/* eslint-disable @typescript-eslint/no-explicit-any */
import type ExcelJSNS from 'exceljs';
import { FErr, isErr, asErr, type Scalar } from '../engine/values';
import { SHORT_DATE } from '../format/numfmt';
import { cellKey, colName, keyCol, keyRow, MAX_COLS, MAX_ROWS, parseRange, rangeName, type Range } from '../model/address';
import { Workbook, type Sheet } from '../model/workbook';
import type { Border, BorderStyle, CellStyle, CondFormat, CompareOp, Validation, ValidationType } from '../model/types';
import type { SheetDoc } from '../doc';
import { fromExcelFormula, toExcelFormula } from './excelFormula';
import { injectXlsxCharts, readXlsxCharts, type ChartToWrite } from './xlsxCharts';
import { resolveSource } from '../charts/chartConfig';

type ExcelJS = typeof ExcelJSNS;
let mod: Promise<ExcelJS> | null = null;
async function excel(): Promise<ExcelJS> {
  mod ??= import('exceljs').then((m: any) => (m.default ?? m) as ExcelJS);
  return mod;
}

/* ================================================================ colours */

const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

function themeColors(wb: any): string[] {
  const xml: string | undefined = wb?._themes?.theme1 ?? wb?.model?.themes?.theme1;
  if (!xml || typeof xml !== 'string') return DEFAULT_THEME;
  const scheme = /<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(xml)?.[1];
  if (!scheme) return DEFAULT_THEME;
  const get = (tag: string) => {
    const block = new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`).exec(scheme)?.[1] ?? '';
    return /srgbClr val="([0-9A-Fa-f]{6})"/.exec(block)?.[1] ?? /lastClr="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
  };
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return order.map((t, i) => (get(t) ?? DEFAULT_THEME[i]).toUpperCase());
}

function applyTint(hex: string, tint: number): string {
  const n = parseInt(hex, 16);
  let r = (n >> 16) & 255;
  let g = (n >> 8) & 255;
  let b = n & 255;
  // HSL lightness adjustment as Excel does
  const max = Math.max(r, g, b) / 255;
  const min = Math.min(r, g, b) / 255;
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    const rr = r / 255;
    const gg = g / 255;
    const bb = b / 255;
    if (max === rr) h = (gg - bb) / d + (gg < bb ? 6 : 0);
    else if (max === gg) h = (bb - rr) / d + 2;
    else h = (rr - gg) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  if (s === 0) r = g = b = Math.round(l * 255);
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = Math.round(hue(p, q, h + 1 / 3) * 255);
    g = Math.round(hue(p, q, h) * 255);
    b = Math.round(hue(p, q, h - 1 / 3) * 255);
  }
  return [r, g, b].map((x) => Math.max(0, Math.min(255, x)).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function colorOf(c: any, theme: string[]): string | undefined {
  if (!c) return undefined;
  let hex: string | undefined;
  if (typeof c.argb === 'string') {
    const a = c.argb.length === 8 ? c.argb.slice(0, 2) : 'FF';
    if (a === '00' && c.argb.length === 8 && c.argb !== '00000000') hex = c.argb.slice(2);
    else hex = c.argb.length === 8 ? c.argb.slice(2) : c.argb;
  } else if (typeof c.theme === 'number') hex = theme[c.theme] ?? undefined;
  else if (typeof c.indexed === 'number') {
    if (c.indexed === 64) return undefined; // system foreground (automatic)
    hex = INDEXED[c.indexed];
  }
  if (!hex) return undefined;
  if (typeof c.tint === 'number' && c.tint) hex = applyTint(hex, c.tint);
  return '#' + hex.toLowerCase();
}

const argb = (hex: string | undefined) => (hex ? { argb: 'FF' + hex.replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6) } : undefined);

/* ============================================================== styles in */

const BORDER_STYLES = new Set<BorderStyle>(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']);

function borderIn(b: any, theme: string[]): Border | undefined {
  if (!b || !b.style) return undefined;
  const style = (BORDER_STYLES.has(b.style) ? b.style : 'thin') as BorderStyle;
  const color = colorOf(b.color, theme);
  return color && color !== '#000000' ? { style, color } : { style };
}

function normalizeNumFmt(fmt: string | undefined): string | undefined {
  if (!fmt || /^general$/i.test(fmt)) return undefined;
  if (fmt === 'mm-dd-yy' || fmt === 'm/d/yy' || fmt === 'm/d/yyyy') return SHORT_DATE;
  return fmt.replace(/\\ /g, ' ');
}

function styleIn(cell: any, theme: string[]): CellStyle {
  const st: CellStyle = {};
  const f = cell.font;
  if (f) {
    if (f.name && f.name !== 'Calibri') st.font = f.name;
    if (f.size && f.size !== 11) st.size = f.size;
    if (f.bold) st.bold = true;
    if (f.italic) st.italic = true;
    if (f.underline) st.underline = f.underline === 'double' || f.underline === 'doubleAccounting' ? 'double' : 'single';
    if (f.strike) st.strike = true;
    const col = colorOf(f.color, theme);
    if (col && col !== '#000000') st.color = col;
  }
  const fill = cell.fill;
  if (fill) {
    if (fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
      const c = colorOf(fill.fgColor, theme) ?? colorOf(fill.bgColor, theme);
      if (c) st.fill = c;
    } else if (fill.type === 'gradient' && fill.stops?.length) {
      const c = colorOf(fill.stops[0].color, theme);
      if (c) st.fill = c;
    }
  }
  const a = cell.alignment;
  if (a) {
    if (a.horizontal && a.horizontal !== 'general') st.hAlign = a.horizontal;
    if (a.vertical && a.vertical !== 'bottom') st.vAlign = a.vertical === 'middle' || a.vertical === 'center' ? 'middle' : a.vertical === 'top' ? 'top' : undefined;
    if (a.wrapText) st.wrap = true;
    if (a.shrinkToFit) st.shrink = true;
    if (a.indent) st.indent = a.indent;
    if (a.textRotation === 'vertical') st.rotation = 255;
    else if (typeof a.textRotation === 'number' && a.textRotation) st.rotation = a.textRotation > 90 ? 90 - a.textRotation : a.textRotation;
  }
  const b = cell.border;
  if (b) {
    const t = borderIn(b.top, theme);
    const bb = borderIn(b.bottom, theme);
    const l = borderIn(b.left, theme);
    const r = borderIn(b.right, theme);
    if (t) st.bt = t;
    if (bb) st.bb = bb;
    if (l) st.bl = l;
    if (r) st.br = r;
  }
  const nf = normalizeNumFmt(cell.numFmt);
  if (nf) st.numFmt = nf;
  const p = cell.protection;
  if (p) {
    if (p.locked === false) st.locked = false;
    if (p.hidden) st.hideFormula = true;
  }
  return st;
}

/* ================================================================ values */

const EPOCH_OFFSET = 25569;
const dateToSerial = (d: Date) => d.getTime() / 86400000 + EPOCH_OFFSET;

function scalarOf(v: any): Scalar | undefined {
  if (v === null || v === undefined) return undefined;
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  if (v instanceof Date) return dateToSerial(v);
  if (typeof v === 'object' && 'error' in v) return new FErr(v.error);
  if (typeof v === 'object' && v.richText) return v.richText.map((r: any) => r.text).join('');
  if (typeof v === 'object' && 'text' in v) return scalarOf(v.text);
  return undefined;
}

/* ================================================================ import */

const OP_IN: Record<string, CompareOp> = {
  between: 'between',
  notBetween: 'notBetween',
  equal: 'equal',
  notEqual: 'notEqual',
  greaterThan: 'greater',
  lessThan: 'less',
  greaterThanOrEqual: 'greaterEqual',
  lessThanOrEqual: 'lessEqual',
};

function formulaOperand(f: string | number | undefined): string {
  if (f === undefined || f === null) return '';
  const s = String(f);
  if (/^-?\d+(\.\d+)?$/.test(s)) return s;
  if (/^".*"$/.test(s)) return s.slice(1, -1).replace(/""/g, '"');
  return '=' + fromExcelFormula(s);
}

function cfStyleIn(style: any, theme: string[]): CondFormat['style'] {
  if (!style) return undefined;
  const out: NonNullable<CondFormat['style']> = {};
  const f = style.font;
  if (f) {
    if (f.bold) out.bold = true;
    if (f.italic) out.italic = true;
    if (f.underline) out.underline = 'single';
    if (f.strike) out.strike = true;
    const c = colorOf(f.color, theme);
    if (c) out.color = c;
  }
  const fill = style.fill;
  if (fill) {
    const c = colorOf(fill.bgColor, theme) ?? colorOf(fill.fgColor, theme);
    if (c) out.fill = c;
  }
  if (style.numFmt) out.numFmt = typeof style.numFmt === 'string' ? style.numFmt : style.numFmt.formatCode;
  return out;
}

function cfIn(entry: any, theme: string[], seq: () => string): CondFormat[] {
  const ranges = String(entry.ref ?? '')
    .split(/\s+/)
    .map((r) => parseRange(r))
    .filter((r): r is Range => !!r);
  if (!ranges.length) return [];
  const out: CondFormat[] = [];
  const rules = [...(entry.rules ?? [])].sort((a: any, b: any) => (a.priority ?? 0) - (b.priority ?? 0));
  for (const r of rules) {
    const base = { id: seq(), ranges, style: cfStyleIn(r.style, theme), stop: !!r.stopIfTrue };
    switch (r.type) {
      case 'cellIs':
        out.push({ ...base, type: 'cell', op: OP_IN[r.operator] ?? 'equal', values: (r.formulae ?? []).map(formulaOperand) });
        break;
      case 'containsText': {
        const t = r.operator === 'notContains' ? 'notText' : r.operator === 'beginsWith' ? 'begins' : r.operator === 'endsWith' ? 'ends' : 'text';
        out.push({ ...base, type: t, text: r.text ?? '' });
        break;
      }
      case 'top10':
        out.push({ ...base, type: 'top', rank: r.rank ?? 10, percent: !!r.percent, bottom: !!r.bottom });
        break;
      case 'aboveAverage':
        out.push({ ...base, type: 'average', above: r.aboveAverage !== false, equalAverage: !!r.equalAverage });
        break;
      case 'duplicateValues':
        out.push({ ...base, type: 'duplicate' });
        break;
      case 'uniqueValues':
        out.push({ ...base, type: 'unique' });
        break;
      case 'containsBlanks':
        out.push({ ...base, type: 'blank' });
        break;
      case 'notContainsBlanks':
        out.push({ ...base, type: 'notBlank' });
        break;
      case 'containsErrors':
        out.push({ ...base, type: 'error' });
        break;
      case 'notContainsErrors':
        out.push({ ...base, type: 'notError' });
        break;
      case 'timePeriod':
        out.push({ ...base, type: 'date', datePeriod: r.timePeriod });
        break;
      case 'expression':
        out.push({ ...base, type: 'formula', values: ['=' + fromExcelFormula(String(r.formulae?.[0] ?? 'FALSE'))] });
        break;
      case 'colorScale':
        out.push({ ...base, type: 'colorScale', style: undefined, colors: (r.color ?? []).map((c: any) => colorOf(c, theme) ?? '#ffffff') });
        break;
      case 'dataBar':
        out.push({ ...base, type: 'dataBar', style: undefined, colors: [colorOf(r.color, theme) ?? '#638ec6'], showValue: r.showValue !== false });
        break;
      case 'iconSet': {
        const known = ['3Arrows', '3TrafficLights', '3Symbols', '3Stars', '4Arrows', '5Arrows', '3Flags', '5Ratings'];
        const name = String(r.iconSet ?? '3TrafficLights1');
        const set = (known.find((k) => name.startsWith(k)) ?? (name.startsWith('3') ? '3TrafficLights' : name.startsWith('4') ? '4Arrows' : '5Arrows')) as CondFormat['iconSet'];
        out.push({ ...base, type: 'iconSet', style: undefined, iconSet: set, showValue: r.showValue !== false });
        break;
      }
    }
  }
  return out;
}

function dvIn(key: string, dv: any, seq: () => string): Validation | null {
  const ranges = key
    .split(/\s+/)
    .map((k) => parseRange(k))
    .filter((r): r is Range => !!r);
  if (!ranges.length) return null;
  const type = (dv.type === 'any' ? 'any' : dv.type) as ValidationType;
  let values: string[] = [];
  if (type === 'list') {
    const f = String(dv.formulae?.[0] ?? '');
    if (/^".*"$/.test(f)) values = f.slice(1, -1).split(',').map((s) => s.trim());
    else values = ['=' + fromExcelFormula(f)];
  } else values = (dv.formulae ?? []).map((f: any) => (f instanceof Date ? String(dateToSerial(f)) : formulaOperand(f)));
  return {
    id: seq(),
    ranges,
    type,
    op: dv.operator ? OP_IN[dv.operator] : undefined,
    values,
    allowBlank: dv.allowBlank !== false,
    showDropdown: true,
    errorStyle: dv.errorStyle ?? 'stop',
    errorTitle: dv.errorTitle || undefined,
    error: dv.error || undefined,
    promptTitle: dv.promptTitle || undefined,
    prompt: dv.prompt || undefined,
  };
}

export async function importXlsx(bytes: Uint8Array): Promise<{ wb: Workbook; warning?: string }> {
  const X = await excel();
  const xwb = new X.Workbook();
  await xwb.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  const theme = themeColors(xwb);
  const wb = new Workbook();
  let n = 0;
  const seq = () => `x${(n++).toString(36)}`;
  const charts = readXlsxCharts(bytes);
  const warnings: string[] = [];
  const imagesById = new Map<number, { src: string }>();
  const spillAreas: { sheet: Sheet; anchor: number; area: Range }[] = [];
  for (const media of (xwb as any).model?.media ?? []) {
    const ext = String(media.extension ?? 'png').toLowerCase();
    const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : ext === 'svg' ? 'image/svg+xml' : ext === 'emf' || ext === 'wmf' ? '' : `image/${ext}`;
    if (!mime) continue;
    const buf: Uint8Array | undefined = media.buffer ? new Uint8Array(media.buffer) : undefined;
    const b64 = media.base64 ?? (buf ? toBase64(buf) : '');
    if (b64) imagesById.set(media.index ?? imagesById.size, { src: `data:${mime};base64,${b64}` });
  }

  xwb.worksheets.forEach((ws: any, index: number) => {
    const sheet = wb.addSheet(wb.uniqueSheetName(ws.name || `Sheet${index + 1}`));
    sheet.hidden = ws.state === 'hidden' || ws.state === 'veryHidden';
    const props = ws.properties ?? {};
    if (props.tabColor) sheet.tabColor = colorOf(props.tabColor, theme);
    const dcw = props.defaultColWidth ?? 8.43;
    sheet.defaultColWidth = Math.round(dcw * 7 + 5);
    if (props.defaultRowHeight) sheet.defaultRowHeight = Math.round((props.defaultRowHeight * 96) / 72);
    else sheet.defaultRowHeight = 20;
    const view = ws.views?.[0];
    if (view) {
      if (view.state === 'frozen') {
        sheet.view.freezeRows = view.ySplit ?? 0;
        sheet.view.freezeCols = view.xSplit ?? 0;
      }
      if (view.showGridLines === false) sheet.view.showGrid = false;
      if (view.zoomScale) sheet.view.zoom = Math.max(0.25, Math.min(4, view.zoomScale / 100));
      if (view.rightToLeft) sheet.view.rtl = true;
    }
    // columns
    (ws.columns ?? []).forEach((col: any, i: number) => {
      if (!col) return;
      const info: { w?: number; hidden?: boolean; s?: number } = {};
      if (col.width !== undefined && col.width !== null && Math.abs(col.width - dcw) > 0.01) info.w = Math.round(col.width * 7 + 5);
      if (col.hidden) info.hidden = true;
      const cs = col.style ? styleIn(col.style, theme) : {};
      if (Object.keys(cs).length) info.s = wb.styleId(cs);
      if (Object.keys(info).length) sheet.cols.set(i, info);
    });
    // rows & cells
    let cellCount = 0;
    ws.eachRow({ includeEmpty: true }, (row: any, rowNumber: number) => {
      const r = rowNumber - 1;
      const info: { h?: number; hidden?: boolean; custom?: boolean; s?: number } = {};
      if (row.height) {
        info.h = Math.round((row.height * 96) / 72);
        info.custom = true;
      }
      if (row.hidden) info.hidden = true;
      if (Object.keys(info).length) sheet.rows.set(r, info);
      row.eachCell({ includeEmpty: true }, (cell: any, colNumber: number) => {
        const c = colNumber - 1;
        if (++cellCount > 3_000_000) return;
        const st = styleIn(cell, theme);
        const out: { v?: Scalar; f?: string; s?: number; link?: string; note?: string } = {};
        const sid = Object.keys(st).length ? wb.styleId(st) : 0;
        if (sid) out.s = sid;
        const model = cell.model ?? {};
        // cells covered by a merge report the merge's value; they only keep their own formatting
        const covered = cell.isMerged && cell.master && cell.master.address !== cell.address;
        const formula: string | undefined = covered ? undefined : (cell.formula ?? model.formula);
        if (covered) {
          // formatting only
        } else if (formula) {
          out.f = fromExcelFormula(String(formula));
          const res = scalarOf(model.result ?? cell.result);
          if (res !== undefined) out.v = res;
          if (model.shareType === 'array' && model.ref) {
            const ar = parseRange(model.ref);
            if (ar && (ar.r2 > ar.r1 || ar.c2 > ar.c1)) spillAreas.push({ sheet, anchor: cellKey(r, c), area: ar });
          }
        } else {
          const v = cell.value;
          if (v && typeof v === 'object' && 'hyperlink' in v) {
            out.link = String(v.hyperlink);
            out.v = scalarOf(v.text);
          } else out.v = scalarOf(v);
        }
        if (cell.hyperlink && !out.link) out.link = cell.hyperlink;
        const note = cell.note;
        if (note) out.note = typeof note === 'string' ? note : (note.texts ?? []).map((t: any) => t.text).join('');
        if (out.v === undefined) delete out.v;
        if (out.v !== undefined || out.f !== undefined || out.s || out.link || out.note) sheet.put(cellKey(r, c), out);
      });
    });
    if (cellCount > 3_000_000) warnings.push(`“${sheet.name}” is very large; only the first 3 million cells were loaded.`);
    // merges
    const merges: string[] = ws.model?.merges ?? [];
    for (const m of merges) {
      const rg = parseRange(m);
      if (rg) sheet.merges.push(rg);
    }
    // conditional formatting
    for (const cf of ws.conditionalFormattings ?? []) sheet.cf.push(...cfIn(cf, theme, seq));
    // data validation
    const dvs = ws.dataValidations?.model ?? {};
    for (const [key, dv] of Object.entries(dvs)) {
      const v = dvIn(key, dv, seq);
      if (v) sheet.validations.push(v);
    }
    // merge validations with identical rules (ExcelJS splits them per cell)
    sheet.validations = mergeValidations(sheet.validations);
    // auto filter
    if (ws.autoFilter) {
      const af = ws.autoFilter;
      const rg = typeof af === 'string' ? parseRange(af) : af.from ? { r1: af.from.row - 1, c1: af.from.column - 1, r2: af.to.row - 1, c2: af.to.column - 1 } : null;
      if (rg) {
        const ext = sheet.extent();
        sheet.filter = { range: { ...rg, r2: Math.max(rg.r2, ext.rows) }, columns: {} };
      }
    }
    // protection
    if (ws.sheetProtection?.sheet) sheet.protection = { enabled: true };
    // images
    for (const img of ws.getImages?.() ?? []) {
      const media = imagesById.get(Number(img.imageId));
      if (!media) continue;
      const tl = img.range?.tl;
      if (!tl) continue;
      const colPx = (col: number, off: number) => Math.floor(col) * sheet.defaultColWidth + (off ?? 0) / 9525;
      const rowPx = (row: number, off: number) => Math.floor(row) * sheet.defaultRowHeight + (off ?? 0) / 9525;
      let w = 200;
      let h = 150;
      if (img.range.ext) {
        w = img.range.ext.width;
        h = img.range.ext.height;
      } else if (img.range.br) {
        w = colPx(img.range.br.nativeCol ?? img.range.br.col, img.range.br.nativeColOff) - colPx(tl.nativeCol ?? tl.col, tl.nativeColOff);
        h = rowPx(img.range.br.nativeRow ?? img.range.br.row, img.range.br.nativeRowOff) - rowPx(tl.nativeRow ?? tl.row, tl.nativeRowOff);
      }
      sheet.images.push({ id: seq(), src: media.src, anchor: { r: Math.floor(tl.nativeRow ?? tl.row), c: Math.floor(tl.nativeCol ?? tl.col), dx: (tl.nativeColOff ?? 0) / 9525, dy: (tl.nativeRowOff ?? 0) / 9525 }, w: Math.max(10, Math.round(w)), h: Math.max(10, Math.round(h)) });
    }
    const sc = charts.get(index);
    if (sc) sheet.charts.push(...sc);
  });

  // spilled array formulas: the other cells only hold cached values
  for (const { sheet, anchor, area } of spillAreas) {
    for (let r = area.r1; r <= area.r2; r++)
      for (let c = area.c1; c <= area.c2; c++) {
        const k = cellKey(r, c);
        if (k === anchor) continue;
        const cell = sheet.cells.get(k);
        if (cell && cell.f === undefined) {
          delete cell.v;
          if (!cell.s && !cell.link && !cell.note) sheet.cells.delete(k);
        }
      }
  }

  // defined names
  for (const dn of (xwb as any).definedNames?.model ?? []) {
    const name: string = dn.name;
    if (!name || name.startsWith('_xlnm.')) continue;
    const ref = (dn.ranges ?? []).join(',');
    if (!ref) continue;
    wb.names.push({ name, ref: fromExcelFormula(ref), sheet: typeof dn.localSheetId === 'number' ? wb.sheets[dn.localSheetId]?.id : undefined });
  }
  wb.props = { title: xwb.title || undefined, author: xwb.creator || undefined, company: xwb.company || undefined, created: xwb.created ? new Date(xwb.created).getTime() : undefined };
  const active = (xwb as any).views?.[0]?.activeTab;
  if (typeof active === 'number') wb.activeSheet = Math.min(active, wb.sheets.length - 1);
  if (!wb.sheets.length) wb.addSheet();
  return { wb, warning: warnings.join(' ') || undefined };
}

function mergeValidations(list: Validation[]): Validation[] {
  const out: Validation[] = [];
  for (const v of list) {
    const key = JSON.stringify({ ...v, id: '', ranges: [] });
    const hit = out.find((o) => JSON.stringify({ ...o, id: '', ranges: [] }) === key);
    if (hit) hit.ranges.push(...v.ranges);
    else out.push(v);
  }
  return out;
}

function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/* ================================================================ export */

const OP_OUT: Record<CompareOp, string> = {
  between: 'between',
  notBetween: 'notBetween',
  equal: 'equal',
  notEqual: 'notEqual',
  greater: 'greaterThan',
  less: 'lessThan',
  greaterEqual: 'greaterThanOrEqual',
  lessEqual: 'lessThanOrEqual',
};

function operandOut(v: string | undefined): string {
  if (v === undefined || v === '') return '0';
  if (v.startsWith('=')) return toExcelFormula(v.slice(1));
  if (/^-?\d+(\.\d+)?$/.test(v.trim())) return v.trim();
  return `"${v.replace(/"/g, '""')}"`;
}

function styleOut(st: CellStyle): any {
  const out: any = {};
  const font: any = {};
  if (st.font) font.name = st.font;
  if (st.size) font.size = st.size;
  if (st.bold) font.bold = true;
  if (st.italic) font.italic = true;
  if (st.underline) font.underline = st.underline === 'double' ? 'double' : true;
  if (st.strike) font.strike = true;
  if (st.color) font.color = argb(st.color);
  if (Object.keys(font).length) out.font = { name: 'Calibri', size: 11, ...font };
  if (st.fill) out.fill = { type: 'pattern', pattern: 'solid', fgColor: argb(st.fill), bgColor: argb(st.fill) };
  const al: any = {};
  if (st.hAlign) al.horizontal = st.hAlign;
  if (st.vAlign) al.vertical = st.vAlign;
  if (st.wrap) al.wrapText = true;
  if (st.shrink) al.shrinkToFit = true;
  if (st.indent) al.indent = st.indent;
  if (st.rotation) al.textRotation = st.rotation === 255 ? 'vertical' : st.rotation < 0 ? 90 - st.rotation : st.rotation;
  if (Object.keys(al).length) out.alignment = al;
  const side = (b?: Border) => (b ? { style: b.style, color: b.color ? argb(b.color) : { argb: 'FF000000' } } : undefined);
  if (st.bt || st.bb || st.bl || st.br) out.border = { top: side(st.bt), bottom: side(st.bb), left: side(st.bl), right: side(st.br) };
  if (st.numFmt) out.numFmt = st.numFmt;
  if (st.locked === false || st.hideFormula) out.protection = { locked: st.locked !== false, hidden: !!st.hideFormula };
  return out;
}

function valueOut(v: Scalar | undefined): any {
  if (v === undefined || v === null) return null;
  if (isErr(v)) return { error: asErr(v).error === '#CIRC!' ? '#REF!' : asErr(v).error };
  return v;
}

function cfOut(cf: CondFormat): any | null {
  const style: any = {};
  if (cf.style) {
    const s = cf.style;
    const font: any = {};
    if (s.bold) font.bold = true;
    if (s.italic) font.italic = true;
    if (s.underline) font.underline = true;
    if (s.strike) font.strike = true;
    if (s.color) font.color = argb(s.color);
    if (Object.keys(font).length) style.font = font;
    if (s.fill) style.fill = { type: 'pattern', pattern: 'solid', bgColor: argb(s.fill) };
    if (s.numFmt) style.numFmt = s.numFmt;
  }
  const base = { priority: 1, style, stopIfTrue: cf.stop || undefined };
  const first = cf.ranges[0];
  const tl = `${colName(first.c1)}${first.r1 + 1}`;
  switch (cf.type) {
    case 'cell':
      return { ...base, type: 'cellIs', operator: OP_OUT[cf.op ?? 'equal'], formulae: (cf.values ?? []).slice(0, cf.op === 'between' || cf.op === 'notBetween' ? 2 : 1).map(operandOut) };
    case 'text':
      return { ...base, type: 'containsText', operator: 'containsText', text: cf.text ?? '', formulae: [`NOT(ISERROR(SEARCH("${(cf.text ?? '').replace(/"/g, '""')}",${tl})))`] };
    case 'notText':
      return { ...base, type: 'containsText', operator: 'notContains', text: cf.text ?? '', formulae: [`ISERROR(SEARCH("${(cf.text ?? '').replace(/"/g, '""')}",${tl}))`] };
    case 'begins':
      return { ...base, type: 'containsText', operator: 'beginsWith', text: cf.text ?? '', formulae: [`LEFT(${tl},LEN("${(cf.text ?? '').replace(/"/g, '""')}"))="${(cf.text ?? '').replace(/"/g, '""')}"`] };
    case 'ends':
      return { ...base, type: 'containsText', operator: 'endsWith', text: cf.text ?? '', formulae: [`RIGHT(${tl},LEN("${(cf.text ?? '').replace(/"/g, '""')}"))="${(cf.text ?? '').replace(/"/g, '""')}"`] };
    case 'top':
      return { ...base, type: 'top10', rank: cf.rank ?? 10, percent: !!cf.percent, bottom: !!cf.bottom };
    case 'average':
      return { ...base, type: 'aboveAverage', aboveAverage: cf.above !== false, equalAverage: !!cf.equalAverage };
    case 'duplicate':
      return { ...base, type: 'duplicateValues' };
    case 'unique':
      return { ...base, type: 'uniqueValues' };
    case 'blank':
      return { ...base, type: 'containsBlanks', formulae: [`LEN(TRIM(${tl}))=0`] };
    case 'notBlank':
      return { ...base, type: 'notContainsBlanks', formulae: [`LEN(TRIM(${tl}))>0`] };
    case 'error':
      return { ...base, type: 'containsErrors', formulae: [`ISERROR(${tl})`] };
    case 'notError':
      return { ...base, type: 'notContainsErrors', formulae: [`NOT(ISERROR(${tl}))`] };
    case 'date':
      return { ...base, type: 'timePeriod', timePeriod: cf.datePeriod ?? 'today' };
    case 'formula':
      return { ...base, type: 'expression', formulae: [toExcelFormula((cf.values?.[0] ?? '=FALSE').replace(/^=/, ''))] };
    case 'colorScale': {
      const cols = cf.colors?.length ? cf.colors : ['#f8696b', '#ffeb84', '#63be7b'];
      const cfvo = cols.length === 3 ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }];
      return { priority: 1, type: 'colorScale', cfvo, color: cols.map((c) => argb(c)) };
    }
    case 'dataBar':
      return { priority: 1, type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: argb(cf.colors?.[0] ?? '#638ec6'), showValue: cf.showValue !== false, gradient: true };
    case 'iconSet': {
      const n = Number((cf.iconSet ?? '3Arrows')[0]);
      const cfvo = Array.from({ length: n }, (_, i) => ({ type: 'percent', value: Math.round((i * 100) / n) }));
      return { priority: 1, type: 'iconSet', iconSet: cf.iconSet === '3TrafficLights' ? '3TrafficLights1' : cf.iconSet ?? '3Arrows', cfvo, showValue: cf.showValue !== false };
    }
  }
  return null;
}

function dvOut(v: Validation): any | null {
  if (v.type === 'any') return null;
  const base: any = {
    allowBlank: v.allowBlank !== false,
    showErrorMessage: true,
    errorStyle: v.errorStyle ?? 'stop',
    errorTitle: v.errorTitle,
    error: v.error,
    showInputMessage: !!v.prompt,
    promptTitle: v.promptTitle,
    prompt: v.prompt,
  };
  if (v.type === 'list') {
    const f = v.values.length === 1 && v.values[0].startsWith('=') ? toExcelFormula(v.values[0].slice(1)) : `"${v.values.join(',').replace(/"/g, '""')}"`;
    return { ...base, type: 'list', formulae: [f] };
  }
  if (v.type === 'checkbox') return { ...base, type: 'list', formulae: ['"TRUE,FALSE"'] };
  if (v.type === 'custom') return { ...base, type: 'custom', formulae: [toExcelFormula((v.values[0] ?? '=TRUE').replace(/^=/, ''))] };
  const two = v.op === 'between' || v.op === 'notBetween' || !v.op;
  return { ...base, type: v.type, operator: OP_OUT[v.op ?? 'between'], formulae: v.values.slice(0, two ? 2 : 1).map(operandOut) };
}

function chartsToWrite(doc: SheetDoc, sheet: Sheet): ChartToWrite[] {
  const out: ChartToWrite[] = [];
  const colPos = (c: number) => {
    let x = 0;
    for (let i = 0; i < c; i++) x += sheet.colWidth(i);
    return x;
  };
  const cellAt = (px: number, axis: 'c' | 'r') => {
    let i = 0;
    let pos = 0;
    const size = (k: number) => (axis === 'c' ? sheet.colWidth(k) : sheet.rowHeight(k));
    while (pos + size(i) <= px && i < (axis === 'c' ? MAX_COLS : MAX_ROWS) - 1) {
      pos += size(i);
      i++;
    }
    return { i, off: px - pos };
  };
  for (const spec of sheet.charts) {
    const src = resolveSource(doc, spec, sheet);
    if (!src) continue;
    const q = /^[A-Za-z_][\w.]*$/.test(src.sheet.name) ? src.sheet.name : `'${src.sheet.name.replace(/'/g, "''")}'`;
    const rows = src.r2 - src.r1 + 1;
    const cols = src.c2 - src.c1 + 1;
    const byRows = spec.seriesInRows ?? cols > rows;
    const v0 = doc.value(src.sheet, src.r1, src.c1 + 1);
    const headerRow = byRows ? typeof doc.value(src.sheet, src.r1 + 1, src.c1) === 'string' : typeof v0 === 'string' || (v0 === undefined && typeof doc.value(src.sheet, src.r1, src.c1) !== 'number');
    const firstColText = byRows ? typeof doc.value(src.sheet, src.r1, src.c1 + 1) === 'string' : typeof doc.value(src.sheet, src.r1 + (headerRow ? 1 : 0), src.c1) === 'string' || spec.type === 'scatter';
    const ref = (r1: number, c1: number, r2: number, c2: number) => `${q}!$${colName(c1)}$${r1 + 1}${r1 === r2 && c1 === c2 ? '' : `:$${colName(c2)}$${r2 + 1}`}`;
    const series: ChartToWrite['series'] = [];
    if (!byRows) {
      const dr1 = src.r1 + (headerRow ? 1 : 0);
      for (let c = src.c1 + (firstColText ? 1 : 0); c <= src.c2; c++)
        series.push({ name: headerRow ? ref(src.r1, c, src.r1, c) : undefined, cat: firstColText ? ref(dr1, src.c1, src.r2, src.c1) : undefined, val: ref(dr1, c, src.r2, c) });
    } else {
      const dc1 = src.c1 + (headerRow ? 1 : 0);
      for (let r = src.r1 + (firstColText ? 1 : 0); r <= src.r2; r++)
        series.push({ name: headerRow ? ref(r, src.c1, r, src.c1) : undefined, cat: firstColText ? ref(src.r1, dc1, src.r1, src.c2) : undefined, val: ref(r, dc1, r, src.c2) });
    }
    if (!series.length) continue;
    const x0 = colPos(spec.anchor.c) + spec.anchor.dx;
    let y0 = 0;
    for (let i = 0; i < spec.anchor.r; i++) y0 += sheet.rowHeight(i);
    y0 += spec.anchor.dy;
    const end = { c: cellAt(x0 + spec.w, 'c'), r: cellAt(y0 + spec.h, 'r') };
    out.push({ spec, series, from: { c: spec.anchor.c, r: spec.anchor.r, dx: spec.anchor.dx, dy: spec.anchor.dy }, to: { c: end.c.i, r: end.r.i, dx: end.c.off, dy: end.r.off } });
  }
  return out;
}

export async function exportXlsx(doc: SheetDoc): Promise<Uint8Array> {
  const X = await excel();
  const src = doc.wb;
  const xwb = new X.Workbook();
  xwb.creator = src.props.author ?? 'Affice';
  xwb.lastModifiedBy = 'Affice';
  xwb.created = src.props.created ? new Date(src.props.created) : new Date();
  xwb.modified = new Date();
  if (src.props.title) xwb.title = src.props.title;
  xwb.calcProperties = { fullCalcOnLoad: true } as any;
  const styleCache = new Map<number, any>();
  const styleFor = (id: number | undefined) => {
    if (!id) return null;
    let s = styleCache.get(id);
    if (!s) {
      s = styleOut(src.style(id));
      styleCache.set(id, s);
    }
    return s;
  };
  const charts = new Map<number, ChartToWrite[]>();

  src.sheets.forEach((sheet, index) => {
    const v = sheet.view;
    const ws = xwb.addWorksheet(sheet.name, {
      state: sheet.hidden ? 'hidden' : 'visible',
      properties: { tabColor: sheet.tabColor ? argb(sheet.tabColor) : undefined, defaultRowHeight: (sheet.defaultRowHeight * 72) / 96, defaultColWidth: Math.max(1, (sheet.defaultColWidth - 5) / 7) } as any,
      views: [
        {
          state: v.freezeRows || v.freezeCols ? 'frozen' : 'normal',
          xSplit: v.freezeCols || undefined,
          ySplit: v.freezeRows || undefined,
          showGridLines: v.showGrid,
          zoomScale: Math.round((v.zoom || 1) * 100),
          rightToLeft: v.rtl || undefined,
          activeCell: v.selection ? `${colName(v.selection.active.c)}${v.selection.active.r + 1}` : undefined,
        } as any,
      ],
    });
    // columns
    for (const [c, info] of sheet.cols) {
      if (c >= MAX_COLS) continue;
      const col = ws.getColumn(c + 1);
      if (info.w !== undefined) col.width = Math.max(0, (info.w - 5) / 7);
      if (info.hidden) col.hidden = true;
      const st = styleFor(info.s);
      if (st) col.style = st;
    }
    // rows
    for (const [r, info] of sheet.rows) {
      const row = ws.getRow(r + 1);
      if (info.h !== undefined) row.height = (info.h * 72) / 96;
      if (info.hidden) row.hidden = true;
      const st = styleFor(info.s);
      if (st) (row as any).style = st;
    }
    for (const r of sheet.filterHidden) ws.getRow(r + 1).hidden = true;
    // cells (row-major for stable output)
    const keys = [...sheet.cells.keys()].sort((a, b) => a - b);
    for (const k of keys) {
      const cellData = sheet.cells.get(k)!;
      const r = keyRow(k);
      const c = keyCol(k);
      const cell = ws.getCell(r + 1, c + 1);
      const val = doc.value(sheet, r, c);
      if (cellData.f !== undefined) {
        const area = doc.engine.spillAreaOf(sheet.id, r, c);
        const formula = toExcelFormula(cellData.f);
        const result = valueOut(val ?? undefined);
        if (area && !doc.engine.isSpillBlocked(sheet.id, r, c) && (area.r2 > area.r1 || area.c2 > area.c1)) {
          cell.value = { formula, result, shareType: 'array', ref: rangeName({ r1: area.r1, c1: area.c1, r2: area.r2, c2: area.c2 }) } as any;
          // write the spilled values so other apps show them
          for (let rr = area.r1; rr <= area.r2; rr++)
            for (let cc = area.c1; cc <= area.c2; cc++) {
              if (rr === r && cc === c) continue;
              const sv = doc.value(sheet, rr, cc);
              if (sv !== undefined && sv !== null && !sheet.cells.get(cellKey(rr, cc))) (ws.getCell(rr + 1, cc + 1) as any).value = valueOut(sv);
            }
        } else cell.value = { formula, result } as any;
      } else if (cellData.link) cell.value = { text: String(cellData.v ?? cellData.link), hyperlink: cellData.link } as any;
      else if (cellData.v !== undefined) cell.value = valueOut(cellData.v);
      const st = styleFor(cellData.s);
      if (st) cell.style = st;
      else if (cellData.link) cell.font = { color: { argb: 'FF0563C1' }, underline: true } as any;
      if (cellData.note) cell.note = cellData.note;
    }
    // merges
    for (const m of sheet.merges) {
      try {
        ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
      } catch {
        /* overlapping merge */
      }
    }
    // conditional formatting
    for (const cf of sheet.cf) {
      const rule = cfOut(cf);
      if (!rule) continue;
      ws.addConditionalFormatting({ ref: cf.ranges.map(rangeName).join(' '), rules: [rule] } as any);
    }
    // validation
    for (const dv of sheet.validations) {
      const rule = dvOut(dv);
      if (!rule) continue;
      for (const rg of dv.ranges) (ws as any).dataValidations.add(rangeName(rg), rule);
    }
    // filter
    if (sheet.filter) ws.autoFilter = rangeName(sheet.filter.range);
    // pictures
    for (const img of sheet.images) {
      const m = /^data:image\/(\w+);base64,(.*)$/.exec(img.src);
      if (!m) continue;
      const ext = m[1] === 'jpeg' ? 'jpeg' : m[1] === 'gif' ? 'gif' : 'png';
      const id = xwb.addImage({ base64: m[2], extension: ext as 'png' });
      ws.addImage(id, { tl: { col: img.anchor.c + img.anchor.dx / Math.max(1, sheet.colWidth(img.anchor.c)), row: img.anchor.r + img.anchor.dy / Math.max(1, sheet.rowHeight(img.anchor.r)) } as any, ext: { width: img.w, height: img.h } });
    }
    // protection
    if (sheet.protection?.enabled) {
      (ws as any).sheetProtection = { sheet: true, objects: true, scenarios: true, ...(sheet.protection.password ? { password: sheet.protection.password } : {}) };
    }
    const cw = chartsToWrite(doc, sheet);
    if (cw.length) charts.set(index, cw);
  });

  // defined names
  for (const n of src.names) {
    const ref = toExcelFormula(n.ref.replace(/^=/, ''));
    try {
      const localSheetId = n.sheet !== undefined ? src.sheets.findIndex((s) => s.id === n.sheet) : undefined;
      (xwb as any).definedNames.model = [...((xwb as any).definedNames.model ?? []), { name: n.name, ranges: [ref], ...(localSheetId !== undefined && localSheetId >= 0 ? { localSheetId } : {}) }];
    } catch {
      /* ExcelJS only understands range names */
    }
  }
  xwb.views = [{ x: 0, y: 0, width: 28800, height: 17000, firstSheet: 0, activeTab: src.activeSheet, visibility: 'visible' } as any];
  const buf = new Uint8Array((await xwb.xlsx.writeBuffer()) as ArrayBuffer);
  return injectXlsxCharts(buf, charts);
}

