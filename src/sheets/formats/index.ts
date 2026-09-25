import { gunzipSync, gzipSync, strFromU8, strToU8 } from 'fflate';
import type { TabSource } from '@/app/workspace';
import { encoder, escapeHtml } from '@/lib/utils';
import { isErr, asErr, type Scalar } from '../engine/values';
import { formatValue, parseInput } from '../format/numfmt';
import { cellKey } from '../model/address';
import { Workbook, workbookFromJSON, workbookToJSON, type Sheet, type WorkbookJSON } from '../model/workbook';
import type { SheetDoc } from '../doc';

export interface LoadedWorkbook {
  wb: Workbook;
  warning?: string;
}

/* ================================================================== text */

export function decodeText(bytes: Uint8Array): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return new TextDecoder('utf-8').decode(bytes.subarray(3));
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

/** Picks the most consistent delimiter from the first lines. */
export function sniffDelimiter(text: string): string {
  const lines = text.split(/\r?\n/).slice(0, 20).filter(Boolean);
  let best = ',';
  let bestScore = -1;
  for (const d of [',', ';', '\t', '|']) {
    const counts = lines.map((l) => splitLine(l, d).length);
    if (!counts.length || counts[0] < 2) continue;
    const consistent = counts.filter((c) => c === counts[0]).length;
    const score = consistent * 10 + counts[0];
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

function splitLine(line: string, d: string): string[] {
  return parseDelimited(line, d)[0] ?? [];
}

export function parseDelimited(text: string, d: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = false;
      } else cur += ch;
      continue;
    }
    if (ch === '"' && cur === '') q = true;
    else if (ch === d) {
      row.push(cur);
      cur = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
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

/** Fills a sheet from delimited text (numbers, dates and formulas are recognised). */
export function fillFromRows(wb: Workbook, sheet: Sheet, rows: string[][], r0 = 0, c0 = 0): void {
  rows.forEach((row, i) =>
    row.forEach((t, j) => {
      if (t === '') return;
      const k = cellKey(r0 + i, c0 + j);
      if (t.startsWith('=') && t.length > 1) {
        sheet.put(k, { f: t.slice(1) });
        return;
      }
      // keep leading zeros (IDs, zip codes) as text
      if (/^0\d+$/.test(t.trim())) {
        sheet.put(k, { v: t });
        return;
      }
      const p = parseInput(t);
      const cell: { v?: Scalar; s?: number } = { v: p.value === null ? undefined : p.value };
      if (p.format && typeof p.value === 'number') cell.s = wb.styleId({ numFmt: p.format });
      sheet.put(k, cell);
    }),
  );
  // widen columns to fit short content
  const maxLen: number[] = [];
  rows.slice(0, 500).forEach((row) => row.forEach((t, j) => (maxLen[j] = Math.max(maxLen[j] ?? 0, t.length))));
  maxLen.forEach((n, j) => {
    const w = Math.min(320, Math.max(64, n * 7.5 + 12));
    if (w > sheet.defaultColWidth + 4) sheet.cols.set(c0 + j, { w: Math.round(w) });
  });
}

/* ================================================================ native */

export function serializeAfsheet(doc: SheetDoc): Uint8Array {
  return gzipSync(strToU8(JSON.stringify(workbookToJSON(doc.wb))), { level: 6 });
}

export function parseAfsheet(bytes: Uint8Array): Workbook {
  const text = bytes[0] === 0x1f && bytes[1] === 0x8b ? strFromU8(gunzipSync(bytes)) : decodeText(bytes);
  const j = JSON.parse(text) as WorkbookJSON;
  if (j.format !== 'affice-sheet') throw new Error('This is not an Affice Sheet file.');
  return workbookFromJSON(j);
}

/* ================================================================== load */

export async function loadWorkbook(source: TabSource | undefined): Promise<LoadedWorkbook> {
  if (!source || source.type === 'blank') {
    const wb = new Workbook();
    wb.addSheet();
    return { wb };
  }
  if (source.type === 'template') {
    const { sheetTemplate } = await import('../templates');
    return { wb: sheetTemplate(source.id) };
  }
  if (source.type === 'recovery') return { wb: workbookFromJSON(JSON.parse(source.data)) };
  if (source.type === 'data') return source.data as LoadedWorkbook;
  const { file, formatId } = source;
  const bytes = file.data;
  switch (formatId) {
    case 'xlsx': {
      const { importXlsx } = await import('./xlsx');
      return importXlsx(bytes);
    }
    case 'ods': {
      const { importOds } = await import('./ods');
      return importOds(bytes);
    }
    case 'afsheet':
      return { wb: parseAfsheet(bytes) };
    case 'csv':
    case 'tsv': {
      const text = decodeText(bytes);
      const d = formatId === 'tsv' ? '\t' : sniffDelimiter(text);
      const wb = new Workbook();
      const sheet = wb.addSheet(file.name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1');
      const rows = parseDelimited(text, d);
      fillFromRows(wb, sheet, rows);
      return { wb, warning: rows.length > 1_048_576 ? 'The file has more rows than a sheet can hold; extra rows were skipped.' : undefined };
    }
    default: {
      const { importWithSheetJS } = await import('./sheetjs');
      return importWithSheetJS(bytes);
    }
  }
}

/* ================================================================ export */

function cellText(doc: SheetDoc, sheet: Sheet, r: number, c: number, raw: boolean): string {
  const v = doc.value(sheet, r, c);
  if (v === undefined || v === null) return '';
  if (isErr(v)) return asErr(v).error;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  const fmt = doc.styleOf(sheet, r, c).numFmt;
  if (typeof v === 'number' && (raw || !fmt || /^general$/i.test(fmt))) return String(parseFloat(v.toPrecision(15)));
  return formatValue(v, fmt).text;
}

function usedGrid(doc: SheetDoc, sheet: Sheet, raw: boolean): string[][] {
  const ext = doc.engine.extent(sheet.id);
  const out: string[][] = [];
  for (let r = 0; r <= ext.rows; r++) {
    const row: string[] = [];
    for (let c = 0; c <= ext.cols; c++) row.push(cellText(doc, sheet, r, c, raw));
    out.push(row);
  }
  return out;
}

function delimited(rows: string[][], d: string): string {
  const q = (s: string) => (s.includes(d) || /["\r\n]/.test(s) || /^\s|\s$/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  return rows.map((row) => row.map(q).join(d)).join('\r\n') + '\r\n';
}

function htmlTable(doc: SheetDoc, sheet: Sheet): string {
  const ext = doc.engine.extent(sheet.id);
  const covered = new Set<number>();
  for (const m of sheet.merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(cellKey(r, c));
  let out = '<table>\n';
  for (let r = 0; r <= ext.rows; r++) {
    if (sheet.isRowHidden(r)) continue;
    out += '<tr>';
    for (let c = 0; c <= ext.cols; c++) {
      if (sheet.isColHidden(c) || covered.has(cellKey(r, c))) continue;
      const m = sheet.merges.find((x) => x.r1 === r && x.c1 === c);
      const span = m ? ` rowspan="${m.r2 - m.r1 + 1}" colspan="${m.c2 - m.c1 + 1}"` : '';
      const st = doc.styleOf(sheet, r, c);
      const css = [st.bold ? 'font-weight:bold' : '', st.italic ? 'font-style:italic' : '', st.color ? `color:${st.color}` : '', st.fill ? `background:${st.fill}` : '', st.hAlign && st.hAlign !== 'general' ? `text-align:${st.hAlign}` : typeof doc.value(sheet, r, c) === 'number' ? 'text-align:right' : ''].filter(Boolean).join(';');
      const tag = r === 0 && st.bold ? 'th' : 'td';
      const link = sheet.get(r, c)?.link;
      const text = escapeHtml(cellText(doc, sheet, r, c, false));
      out += `<${tag}${span}${css ? ` style="${css}"` : ''}>${link ? `<a href="${escapeHtml(link)}">${text}</a>` : text}</${tag}>`;
    }
    out += '</tr>\n';
  }
  return out + '</table>';
}

export async function exportWorkbook(doc: SheetDoc, formatId: string, title: string): Promise<Uint8Array> {
  const sheet = doc.sheet;
  switch (formatId) {
    case 'xlsx': {
      const { exportXlsx } = await import('./xlsx');
      return exportXlsx(doc);
    }
    case 'ods': {
      const { exportOds } = await import('./ods');
      return exportOds(doc);
    }
    case 'xls': {
      const { exportWithSheetJS } = await import('./sheetjs');
      return exportWithSheetJS(doc, formatId);
    }
    case 'afsheet':
      return serializeAfsheet(doc);
    case 'csv':
      return encoder.encode('﻿' + delimited(usedGrid(doc, sheet, false), ','));
    case 'tsv':
      return encoder.encode(delimited(usedGrid(doc, sheet, false), '\t'));
    case 'json': {
      const rows = usedGrid(doc, sheet, true);
      const [head, ...body] = rows;
      const keys = (head ?? []).map((h, i) => h || `Column${i + 1}`);
      const data = body.map((_, ri) => {
        const o: Record<string, unknown> = {};
        keys.forEach((k, ci) => {
          const v = doc.value(sheet, ri + 1, ci);
          o[k] = v === undefined ? null : isErr(v) ? asErr(v).error : v;
        });
        return o;
      });
      return encoder.encode(JSON.stringify(data, null, 2));
    }
    case 'md': {
      const rows = usedGrid(doc, sheet, false).map((r) => r.map((t) => t.replace(/\|/g, '\\|').replace(/\n/g, '<br>')));
      if (!rows.length) return encoder.encode('');
      const width = Math.max(...rows.map((r) => r.length));
      const pad = (r: string[]) => [...r, ...Array(width - r.length).fill('')];
      const [head, ...body] = rows;
      const md = [`| ${pad(head).join(' | ')} |`, `| ${Array(width).fill('---').join(' | ')} |`, ...body.map((r) => `| ${pad(r).join(' | ')} |`)].join('\n');
      return encoder.encode(md + '\n');
    }
    case 'html': {
      const body = doc.wb.sheets
        .filter((s) => !s.hidden)
        .map((s) => `<h2>${escapeHtml(s.name)}</h2>\n${htmlTable(doc, s)}`)
        .join('\n');
      return encoder.encode(`<!doctype html>
<html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title>
<style>body{font-family:Calibri,Carlito,system-ui,sans-serif;font-size:11pt;margin:24px;color:#1a1d24}h2{font-size:14pt;margin:24px 0 8px}table{border-collapse:collapse;margin-bottom:24px}td,th{border:1px solid #d9dce3;padding:3px 8px;white-space:pre-wrap}th{background:#f2f4f8;text-align:left}</style>
</head><body>
${body}
</body></html>`);
    }
    default:
      throw new Error(`Unsupported format ${formatId}`);
  }
}

