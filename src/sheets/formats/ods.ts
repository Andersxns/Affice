/**
 * OpenDocument Spreadsheet (.ods) reader and writer.
 *
 * Keeps values, formulas, cell styles, number formats, merges, notes, links, column widths and row
 * heights, conditional formatting, data validation, named ranges, filters, frozen panes, zoom, print
 * settings, protection, pictures and sheet colours. The layout follows what LibreOffice writes, so
 * files open the same in LibreOffice, Excel and other OpenDocument apps.
 */
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { bytesToDataUrl, dataUrlToBytes, escapeXml, extFromMime, mimeFromExt } from '@/lib/utils';
import { errFromCode, isErr, asErr, type Scalar } from '../engine/values';
import { tokenize } from '../engine/parser';
import { formatValue, partsToSerial, serialToParts, SHORT_DATE } from '../format/numfmt';
import { cellKey, colIndex, keyCol, keyRow, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { Workbook, type Sheet } from '../model/workbook';
import {
  DEFAULT_PRINT,
  type AutoFilter,
  type Border,
  type BorderStyle,
  type CellStyle,
  type ColumnFilter,
  type CompareOp,
  type CondFormat,
  type DefinedName,
  type ImageSpec,
  type SheetPrint,
  type Validation,
  type ValidationType,
} from '../model/types';
import type { SheetDoc } from '../doc';
import { fromOdsFormula, odsRangeAddress, odsRefToExcel, odsSheet, splitRangeList, toOdsFormula } from './odsFormula';
import { NS_NUMBER, numFmtToOds, odsToNumFmt, odsValueType } from './odsNumber';
import { odsChartObject, readOdsChart } from './odsCharts';
import { fallback, isFallback, scanRegion, splitContent, type CellAttrs, type CellData, type ColData, type RowData, type SplitContent } from './odsScan';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
  dc: 'http://purl.org/dc/elements/1.1/',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  number: NS_NUMBER,
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  of: 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
  config: 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
  ooo: 'http://openoffice.org/2004/office',
  tableooo: 'http://openoffice.org/2009/table',
  calcext: 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
  loext: 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
};
const NS_MANIFEST = 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0';
const MIME = 'application/vnd.oasis.opendocument.spreadsheet';
const XMLNS = Object.entries(NS)
  .map(([k, v]) => `xmlns:${k}="${v}"`)
  .join(' ');
const XML_DECL = '<?xml version="1.0" encoding="UTF-8"?>\n';

/** Columns written per sheet when fewer are used: older LibreOffice versions stop at 1024. */
const MIN_COLS = 1024;
/** Horizontal indent of one indent level, in px. */
const INDENT_PX = 14;
const MAX_CELLS = 3_000_000;

const esc = escapeXml;
const inch = (px: number) => `${Number((px / 96).toFixed(4))}in`;

function lengthPx(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^(-?[\d.]+)\s*(in|cm|mm|pt|pc|px)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  switch (m[2]) {
    case 'in':
      return n * 96;
    case 'cm':
      return (n * 96) / 2.54;
    case 'mm':
      return (n * 96) / 25.4;
    case 'pt':
      return (n * 96) / 72;
    case 'pc':
      return n * 16;
    default:
      return n;
  }
}

/** Style names must be XML names: other characters are written as _hex_ (like LibreOffice). */
function styleNameOf(display: string): string {
  let out = '';
  for (const ch of display) out += /[A-Za-z0-9.-]/.test(ch) ? ch : `_${ch.codePointAt(0)!.toString(16)}_`;
  return /^[A-Za-z_]/.test(out) ? out : `_${out}`;
}

/* ============================================================ values */

/** An Excel serial date as an ODF date value ("2026-09-25" or "2026-09-25T13:30:00"). */
const DAY_MS = 86_400_000;
const ODF_EPOCH = Date.UTC(1899, 11, 30);

/**
 * An Excel serial date as an ODF date value ("2026-09-25" or "2026-09-25T13:30:00"). Serials below 1
 * (times of day) count from 1899-12-30 like in LibreOffice; later ones keep their Excel calendar date.
 */
export function serialToOdsDate(serial: number): string | undefined {
  if (!Number.isFinite(serial) || serial < -657434 || serial > 2958465 || (serial >= 60 && serial < 61)) return undefined;
  const pad = (n: number, w = 2) => String(n).padStart(w, '0');
  let p: { y: number; m: number; d: number; H: number; M: number; S: number; ms: number };
  if (serial < 1) {
    const dt = new Date(ODF_EPOCH + Math.round(serial * DAY_MS));
    p = { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), H: dt.getUTCHours(), M: dt.getUTCMinutes(), S: dt.getUTCSeconds(), ms: dt.getUTCMilliseconds() };
  } else p = serialToParts(serial);
  const date = `${p.y < 0 ? '-' : ''}${pad(Math.abs(p.y), 4)}-${pad(p.m)}-${pad(p.d)}`;
  if (!p.H && !p.M && !p.S && !p.ms) return date;
  return `${date}T${pad(p.H)}:${pad(p.M)}:${pad(p.S)}${p.ms ? `.${pad(Math.round(p.ms), 3)}` : ''}`;
}

export function odsDateToSerial(v: string): number | undefined {
  const m = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(v.trim());
  if (!m) return undefined;
  const [y, mo, d, H, M, S] = [Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0)];
  const day = Date.UTC(y, mo - 1, d);
  // on or before LibreOffice's day 0, count from it (times are stored on 1899-12-30)
  if (day <= ODF_EPOCH) return Math.round((day - ODF_EPOCH) / DAY_MS) + (H * 3600 + M * 60 + S) / 86400;
  return partsToSerial(y, mo, d, H, M, S);
}

/** A day fraction as an ODF duration ("PT13H30M00S"). */
export function serialToOdsTime(v: number): string {
  const neg = v < 0;
  const us = Math.round(Math.abs(v) * 86400 * 1e6);
  const h = Math.floor(us / 3.6e9);
  const m = Math.floor((us - h * 3.6e9) / 6e7);
  const rest = us - h * 3.6e9 - m * 6e7;
  const whole = Math.floor(rest / 1e6);
  const frac = rest - whole * 1e6;
  const pad = (n: number) => (n < 10 ? `0${n}` : String(n));
  return `${neg ? '-' : ''}PT${pad(h)}H${pad(m)}M${pad(whole)}${frac ? `.${String(frac).padStart(6, '0').replace(/0+$/, '')}` : ''}S`;
}

export function odsTimeToSerial(v: string): number | undefined {
  const m = /^(-)?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(v.trim());
  if (!m) return undefined;
  const secs = Number(m[2] ?? 0) * 86400 + Number(m[3] ?? 0) * 3600 + Number(m[4] ?? 0) * 60 + Number(m[5] ?? 0);
  return (m[1] ? -secs : secs) / 86400;
}

/** Cell text as ODF paragraphs: runs of spaces and tabs need their own elements. */
function spanText(s: string): string {
  let out = '';
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === ' ') {
      let j = i;
      while (s[j] === ' ') j++;
      const n = j - i;
      if (i === 0) out += n === 1 ? '<text:s/>' : `<text:s text:c="${n}"/>`;
      else {
        out += ' ';
        if (n === 2) out += '<text:s/>';
        else if (n > 2) out += `<text:s text:c="${n - 1}"/>`;
      }
      i = j;
    } else if (ch === '\t') {
      out += '<text:tab/>';
      i++;
    } else {
      let j = i;
      while (j < s.length && s[j] !== ' ' && s[j] !== '\t') j++;
      out += esc(s.slice(i, j));
      i = j;
    }
  }
  return out;
}

function paragraphs(s: string, link?: string): string {
  return s
    .split(/\r\n|\r|\n/)
    .map((line) => `<text:p>${link ? `<text:a xlink:href="${esc(link)}" xlink:type="simple">${spanText(line)}</text:a>` : spanText(line)}</text:p>`)
    .join('');
}

/* ============================================================ writing: styles */

const BORDER_OUT: Record<BorderStyle, [width: string, kind: string]> = {
  hair: ['0.06pt', 'solid'],
  thin: ['0.74pt', 'solid'],
  medium: ['1.76pt', 'solid'],
  thick: ['2.49pt', 'solid'],
  dashed: ['0.74pt', 'dashed'],
  mediumDashed: ['1.76pt', 'dashed'],
  dotted: ['0.74pt', 'dotted'],
  double: ['2.01pt', 'double'],
  dashDot: ['0.74pt', 'dash-dot'],
  mediumDashDot: ['1.76pt', 'dash-dot'],
  slantDashDot: ['1.76pt', 'dash-dot'],
  dashDotDot: ['0.74pt', 'dash-dot-dot'],
  mediumDashDotDot: ['1.76pt', 'dash-dot-dot'],
};

type StyleProps = Pick<CellStyle, 'font' | 'size' | 'bold' | 'italic' | 'underline' | 'strike' | 'color' | 'fill' | 'bt' | 'br' | 'bb' | 'bl'> & Partial<CellStyle>;

/** The property elements of a cell style. */
function cellStyleBody(st: StyleProps, fonts: Set<string>): string {
  const cell: string[] = [];
  const para: string[] = [];
  const text: string[] = [];
  if (st.fill) cell.push(`fo:background-color="${esc(st.fill)}"`);
  for (const [side, b] of [
    ['top', st.bt],
    ['bottom', st.bb],
    ['left', st.bl],
    ['right', st.br],
  ] as const) {
    if (!b) continue;
    const [w, kind] = BORDER_OUT[b.style] ?? BORDER_OUT.thin;
    cell.push(`fo:border-${side}="${w} ${kind} ${esc(b.color ?? '#000000')}"`);
    if (kind === 'double') cell.push(`style:border-line-width-${side}="0.0071in 0.0102in 0.0071in"`);
  }
  const h = st.hAlign;
  if (h && h !== 'general') {
    cell.push('style:text-align-source="fix"');
    if (h === 'fill') cell.push('style:repeat-content="true"');
    para.push(`fo:text-align="${h === 'left' || h === 'fill' ? 'start' : h === 'right' ? 'end' : h === 'justify' || h === 'distributed' ? 'justify' : 'center'}"`);
  }
  if (st.indent) para.push(`fo:margin-left="${inch(st.indent * INDENT_PX)}"`);
  if (st.wrap) cell.push('fo:wrap-option="wrap"');
  if (st.shrink) cell.push('style:shrink-to-fit="true"');
  if (st.vAlign) cell.push(`style:vertical-align="${st.vAlign}"`);
  if (st.rotation === 255) cell.push('style:direction="ttb"');
  else if (st.rotation) cell.push(`style:rotation-angle="${(Math.round(st.rotation) + 360) % 360}"`);
  if (st.locked === false || st.hideFormula) cell.push(`style:cell-protect="${st.locked === false ? (st.hideFormula ? 'formula-hidden' : 'none') : 'protected formula-hidden'}"`);
  if (st.font) {
    fonts.add(st.font);
    text.push(`style:font-name="${esc(st.font)}"`);
  }
  if (st.size) text.push(`fo:font-size="${st.size}pt" style:font-size-asian="${st.size}pt" style:font-size-complex="${st.size}pt"`);
  if (st.bold) text.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
  if (st.italic) text.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
  if (st.underline) text.push(`style:text-underline-style="solid"${st.underline === 'double' ? ' style:text-underline-type="double"' : ''} style:text-underline-width="auto" style:text-underline-color="font-color"`);
  if (st.strike) text.push('style:text-line-through-style="solid" style:text-line-through-type="single"');
  if (st.color) text.push(`fo:color="${esc(st.color)}"`);
  return (
    (cell.length ? `<style:table-cell-properties ${cell.join(' ')}/>` : '') +
    (para.length ? `<style:paragraph-properties ${para.join(' ')}/>` : '') +
    (text.length ? `<style:text-properties ${text.join(' ')}/>` : '')
  );
}

function fontFaces(fonts: Set<string>): string {
  const generic: Record<string, string> = { Calibri: 'swiss', Arial: 'swiss', Cambria: 'roman', 'Times New Roman': 'roman', 'Courier New': 'modern' };
  return `<office:font-face-decls>${[...fonts]
    .map((f) => `<style:font-face style:name="${esc(f)}" svg:font-family="${esc(/[\s,]/.test(f) ? `'${f}'` : f)}"${generic[f] ? ` style:font-family-generic="${generic[f]}"` : ''}/>`)
    .join('')}</office:font-face-decls>`;
}

/* ============================================================ writing: conditions */

const OP_ODS: Record<CompareOp, string> = { equal: '=', notEqual: '!=', greater: '>', less: '<', greaterEqual: '>=', lessEqual: '<=', between: 'between', notBetween: 'not-between' };

/** A condition operand (number, text or "=formula") as an OpenFormula expression. */
function operandOds(v: string | undefined): string {
  if (v === undefined || v === '') return '0';
  if (v.startsWith('=')) return toOdsFormula(v.slice(1)).slice(4);
  if (/^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(v.trim())) return v.trim();
  return `"${v.replace(/"/g, '""')}"`;
}

const quoted = (s: string | undefined) => `"${(s ?? '').replace(/"/g, '""')}"`;

function validationCondition(v: Validation): string | undefined {
  const a = operandOds(v.values[0]);
  const b = operandOds(v.values[1]);
  const op = v.op ?? 'between';
  const compare = (fn: string) =>
    op === 'between' ? `${fn}-is-between(${a},${b})` : op === 'notBetween' ? `${fn}-is-not-between(${a},${b})` : `${fn}()${OP_ODS[op]}${a}`;
  switch (v.type) {
    case 'list':
      if (v.values.length === 1 && v.values[0].startsWith('=')) return `of:cell-content-is-in-list(${toOdsFormula(v.values[0].slice(1)).slice(4)})`;
      return `of:cell-content-is-in-list(${v.values.map(quoted).join(';')})`;
    case 'checkbox':
      return 'of:cell-content-is-in-list(TRUE();FALSE())';
    case 'custom':
      return `of:is-true-formula(${toOdsFormula((v.values[0] ?? '=TRUE').replace(/^=/, '')).slice(4)})`;
    case 'whole':
      return `of:cell-content-is-whole-number() and ${compare('cell-content')}`;
    case 'decimal':
      return `of:cell-content-is-decimal-number() and ${compare('cell-content')}`;
    case 'date':
      return `of:cell-content-is-date() and ${compare('cell-content')}`;
    case 'time':
      return `of:cell-content-is-time() and ${compare('cell-content')}`;
    case 'textLength':
      return `of:${compare('cell-content-text-length')}`;
    default:
      return undefined;
  }
}

function validationXml(v: Validation, name: string, sheet: Sheet): string | undefined {
  const cond = validationCondition(v);
  if (!cond && !v.prompt) return undefined;
  const first = v.ranges[0];
  const attrs = [`table:name="${name}"`];
  if (cond) attrs.push(`table:condition="${esc(cond)}"`);
  attrs.push(`table:allow-empty-cell="${v.allowBlank !== false}"`);
  if (v.type === 'list' || v.type === 'checkbox') attrs.push(`table:display-list="${v.showDropdown === false ? 'none' : 'unsorted'}"`);
  attrs.push(`table:base-cell-address="${esc(odsRangeAddress(sheet.name, first.r1, first.c1, first.r1, first.c1))}"`);
  let body = '';
  if (v.prompt || v.promptTitle) body += `<table:help-message${v.promptTitle ? ` table:title="${esc(v.promptTitle)}"` : ''} table:display="true">${paragraphs(v.prompt ?? '')}</table:help-message>`;
  if (cond) body += `<table:error-message table:message-type="${v.errorStyle ?? 'stop'}"${v.errorTitle ? ` table:title="${esc(v.errorTitle)}"` : ''} table:display="true">${v.error ? paragraphs(v.error) : ''}</table:error-message>`;
  return `<table:content-validation ${attrs.join(' ')}>${body}</table:content-validation>`;
}

const DATE_PERIOD_OUT: Record<NonNullable<CondFormat['datePeriod']>, string> = {
  yesterday: 'yesterday',
  today: 'today',
  tomorrow: 'tomorrow',
  last7Days: 'last-7-days',
  lastWeek: 'last-week',
  thisWeek: 'this-week',
  nextWeek: 'next-week',
  lastMonth: 'last-month',
  thisMonth: 'this-month',
  nextMonth: 'next-month',
};

const ICON_SETS_OUT: Record<NonNullable<CondFormat['iconSet']>, string> = {
  '3Arrows': '3Arrows',
  '3TrafficLights': '3TrafficLights1',
  '3Symbols': '3Symbols',
  '3Stars': '3Stars',
  '4Arrows': '4Arrows',
  '5Arrows': '5Arrows',
  '3Flags': '3Flags',
  '5Ratings': '5Rating',
};

function cfConditionValue(cf: CondFormat, tl: string): string | undefined {
  const text = quoted(cf.text);
  switch (cf.type) {
    case 'cell': {
      const op = cf.op ?? 'equal';
      const a = operandOds(cf.values?.[0]);
      if (op === 'between' || op === 'notBetween') return `${OP_ODS[op]}(${a},${operandOds(cf.values?.[1])})`;
      return `${OP_ODS[op]}${a}`;
    }
    case 'text':
      return `contains-text(${text})`;
    case 'notText':
      return `not-contains-text(${text})`;
    case 'begins':
      return `begins-with(${text})`;
    case 'ends':
      return `ends-with(${text})`;
    case 'top':
      return `${cf.bottom ? 'bottom' : 'top'}-${cf.percent ? 'percent' : 'elements'}(${cf.rank ?? 10})`;
    case 'average':
      return `${cf.above === false ? 'below' : 'above'}${cf.equalAverage ? '-equal' : ''}-average`;
    case 'duplicate':
      return 'duplicate';
    case 'unique':
      return 'unique';
    case 'blank':
      return `formula-is(LEN(TRIM(${tl}))=0)`;
    case 'notBlank':
      return `formula-is(LEN(TRIM(${tl}))>0)`;
    case 'error':
      return 'is-error';
    case 'notError':
      return 'is-no-error';
    case 'formula':
      return `formula-is(${toOdsFormula((cf.values?.[0] ?? '=FALSE').replace(/^=/, '')).slice(4)})`;
    default:
      return undefined;
  }
}

function cfXml(cf: CondFormat, sheet: Sheet, styleName: string | undefined): string | undefined {
  if (!cf.ranges.length) return undefined;
  const target = cf.ranges.map((r) => odsRangeAddress(sheet.name, r.r1, r.c1, Math.min(r.r2, MAX_ROWS - 1), Math.min(r.c2, MAX_COLS - 1))).join(' ');
  const f = cf.ranges[0];
  const base = odsRangeAddress(sheet.name, f.r1, f.c1, f.r1, f.c1);
  let inner: string;
  switch (cf.type) {
    case 'colorScale': {
      const cols = cf.colors?.length ? cf.colors : ['#f8696b', '#ffeb84', '#63be7b'];
      const entry = (type: string, value: number, color: string) => `<calcext:color-scale-entry calcext:value="${value}" calcext:type="${type}" calcext:color="${esc(color)}"/>`;
      inner = `<calcext:color-scale>${entry('minimum', 0, cols[0])}${cols.length > 2 ? entry('percentile', 50, cols[1]) : ''}${entry('maximum', 0, cols[cols.length - 1])}</calcext:color-scale>`;
      break;
    }
    case 'dataBar':
      inner = `<calcext:data-bar calcext:positive-color="${esc(cf.colors?.[0] ?? '#638ec6')}" calcext:negative-color="#ff0000" calcext:axis-position="automatic" calcext:axis-color="#000000" calcext:min-length="10" calcext:max-length="90" calcext:gradient="true"${cf.showValue === false ? ' calcext:show-value="false"' : ''}><calcext:formatting-entry calcext:value="0" calcext:type="minimum"/><calcext:formatting-entry calcext:value="0" calcext:type="maximum"/></calcext:data-bar>`;
      break;
    case 'iconSet': {
      const set = cf.iconSet ?? '3Arrows';
      const n = Number(set[0]);
      const entries = Array.from({ length: n }, (_, i) => `<calcext:formatting-entry calcext:value="${Math.round((i * 100) / n)}" calcext:type="percent"/>`).join('');
      inner = `<calcext:icon-set calcext:icon-set-type="${ICON_SETS_OUT[set]}"${cf.showValue === false ? ' calcext:show-value="false"' : ''}>${entries}</calcext:icon-set>`;
      break;
    }
    case 'date':
      inner = `<calcext:date-is calcext:date="${DATE_PERIOD_OUT[cf.datePeriod ?? 'today']}" calcext:style="${esc(styleName ?? 'Default')}"/>`;
      break;
    default: {
      const value = cfConditionValue(cf, `[.${colLetters(f.c1)}${f.r1 + 1}]`);
      if (!value) return undefined;
      inner = `<calcext:condition calcext:apply-style-name="${esc(styleName ?? 'Default')}" calcext:value="${esc(value)}" calcext:base-cell-address="${esc(base)}"/>`;
    }
  }
  return `<calcext:conditional-format calcext:target-range-address="${esc(target)}">${inner}</calcext:conditional-format>`;
}

function colLetters(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* ============================================================ writing: filters and names */

function filterXml(f: AutoFilter): string {
  const conds: string[] = [];
  const cond = (field: number, op: string, value: string, extra = '') =>
    `<table:filter-condition table:field-number="${field}" table:value="${esc(value)}" table:operator="${esc(op)}" table:data-type="${/^-?\d+(\.\d+)?$/.test(value.trim()) ? 'number' : 'text'}"${extra ? `>${extra}</table:filter-condition>` : '/>'}`;
  for (const [key, cf] of Object.entries(f.columns) as [string, ColumnFilter][]) {
    const field = Number(key) - f.range.c1;
    if (field < 0) continue;
    if (cf.values) {
      const items = [...cf.values, ...(cf.blanks ? [''] : [])];
      if (!items.length) continue;
      conds.push(cond(field, '=', items[0], items.map((v) => `<table:filter-set-item table:value="${esc(v)}"/>`).join('')));
    } else if (cf.condition) {
      const { op, value, value2 } = cf.condition;
      const simple: Record<string, string> = { equal: '=', notEqual: '!=', greater: '>', less: '<', greaterEqual: '>=', lessEqual: '<=', contains: 'contains', notContains: '!contains', begins: 'begins', ends: 'ends' };
      if (op === 'between') conds.push(cond(field, '>=', value), cond(field, '<=', value2 ?? value));
      else if (op === 'notBetween') conds.push(`<table:filter-or>${cond(field, '<', value)}${cond(field, '>', value2 ?? value)}</table:filter-or>`);
      else if (simple[op]) conds.push(cond(field, simple[op], value));
    } else if (cf.top) conds.push(cond(field, `${cf.top.bottom ? 'bottom' : 'top'} ${cf.top.percent ? 'percent' : 'values'}`, String(cf.top.n)));
  }
  if (!conds.length) return '';
  return `<table:filter>${conds.length === 1 ? conds[0] : `<table:filter-and>${conds.join('')}</table:filter-and>`}</table:filter>`;
}

/** Names that are plain references become named ranges; anything else is a named expression. */
function namedXml(n: DefinedName, baseSheet: string): string {
  const ref = n.ref.replace(/^=/, '');
  const toks = tokenize(ref).filter((t) => t.type !== 'ws');
  const plain =
    (toks.length === 1 && (toks[0].type === 'ref' || toks[0].type === 'colrange' || toks[0].type === 'rowrange')) ||
    (toks.length === 3 && toks[0].type === 'ref' && toks[1].type === 'colon' && toks[2].type === 'ref');
  const base = `${odsSheet(baseSheet)}.$A$1`;
  const f = toOdsFormula(ref);
  if (plain) return `<table:named-range table:name="${esc(n.name)}" table:base-cell-address="${esc(base)}" table:cell-range-address="${esc(f.slice(5, -1))}"/>`;
  return `<table:named-expression table:name="${esc(n.name)}" table:base-cell-address="${esc(base)}" table:expression="${esc(f)}"/>`;
}

/* ============================================================ writing: pages */

const PAPER_IN: Record<SheetPrint['paper'], [number, number]> = {
  letter: [8.5, 11],
  a4: [8.2681, 11.6929],
  legal: [8.5, 14],
  a3: [11.6929, 16.5354],
  a5: [5.8268, 8.2681],
  tabloid: [11, 17],
};
const MARGINS_IN: Record<SheetPrint['margins'], [top: number, right: number, bottom: number, left: number]> = {
  normal: [0.75, 0.7, 0.75, 0.7],
  narrow: [0.5, 0.25, 0.5, 0.25],
  wide: [1, 1, 1, 1],
};

/** Header/footer text with Excel-style codes (&P, &N, &D, &T, &A, &F) as ODF fields. */
function headerFooterXml(text: string): string {
  let out = '';
  const parts = text.split(/(&[PNDTAF])/);
  for (const p of parts) {
    switch (p) {
      case '&P':
        out += '<text:page-number>1</text:page-number>';
        break;
      case '&N':
        out += '<text:page-count>1</text:page-count>';
        break;
      case '&D':
        out += '<text:date/>';
        break;
      case '&T':
        out += '<text:time/>';
        break;
      case '&A':
        out += '<text:sheet-name>???</text:sheet-name>';
        break;
      case '&F':
        out += '<text:title>???</text:title>';
        break;
      default:
        out += spanText(p);
    }
  }
  return `<style:region-center><text:p>${out}</text:p></style:region-center>`;
}

function pageLayoutXml(name: string, pr: SheetPrint): string {
  const [pw, ph] = PAPER_IN[pr.paper] ?? PAPER_IN.letter;
  const [w, h] = pr.orientation === 'landscape' ? [ph, pw] : [pw, ph];
  const [mt, mr, mb, ml] = MARGINS_IN[pr.margins] ?? MARGINS_IN.normal;
  const print = ['charts', 'drawings', 'objects', 'zero-values', ...(pr.gridlines ? ['grid'] : []), ...(pr.headings ? ['headers'] : [])].join(' ');
  const scale =
    pr.fit === 'page'
      ? 'style:scale-to-pages="1"'
      : pr.fit === 'width'
        ? 'style:scale-to-X="1" style:scale-to-Y="0" loext:scale-to-X="1" loext:scale-to-Y="0"'
        : `style:scale-to="${Math.round(pr.scale || 100)}%"`;
  const hf = (tag: 'header' | 'footer') =>
    `<style:${tag}-style><style:header-footer-properties fo:min-height="0.2953in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-${tag === 'header' ? 'bottom' : 'top'}="0.0984in"/></style:${tag}-style>`;
  return `<style:page-layout style:name="${name}"><style:page-layout-properties fo:page-width="${w}in" fo:page-height="${h}in" style:print-orientation="${pr.orientation}" fo:margin-top="${mt}in" fo:margin-bottom="${mb}in" fo:margin-left="${ml}in" fo:margin-right="${mr}in" style:print-page-order="ttb" style:first-page-number="continue" ${scale}${pr.centerH ? ' style:table-centering="horizontal"' : ''} style:writing-mode="lr-tb" style:print="${print}"/>${hf('header')}${hf('footer')}</style:page-layout>`;
}

function masterPageXml(name: string, display: string | undefined, layout: string, pr: SheetPrint): string {
  const part = (tag: 'header' | 'footer', text: string | undefined) => (text ? `<style:${tag}>${headerFooterXml(text)}</style:${tag}>` : `<style:${tag} style:display="false"/>`);
  return `<style:master-page style:name="${name}"${display ? ` style:display-name="${esc(display)}"` : ''} style:page-layout-name="${layout}">${part('header', pr.header)}<style:header-left style:display="false"/>${part('footer', pr.footer)}<style:footer-left style:display="false"/></style:master-page>`;
}

/* ============================================================ writing: workbook */

/** UTF-8 bytes of many strings, encoded in medium-sized batches. */
function encodeParts(parts: string[]): Uint8Array {
  const chunks: Uint8Array[] = [];
  let batch: string[] = [];
  let len = 0;
  let total = 0;
  const flush = () => {
    if (!batch.length) return;
    const bytes = strToU8(batch.join(''));
    chunks.push(bytes);
    total += bytes.length;
    batch = [];
    len = 0;
  };
  for (const p of parts) {
    batch.push(p);
    len += p.length;
    if (len >= 1 << 20) flush();
  }
  flush();
  if (chunks.length === 1) return chunks[0];
  const out = new Uint8Array(total);
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  return out;
}

interface Frag {
  xml: string;
  n: number;
}

function pushFrag(list: Frag[], xml: string, n: number): void {
  const last = list[list.length - 1];
  if (last && last.xml === xml) last.n += n;
  else list.push({ xml, n });
}

/** Adds a repeat attribute to a fragment written as `<tag attrs…` (self-closing or not). */
function withRepeat(xml: string, attr: string, n: number): string {
  if (n <= 1) return xml;
  const m = /^<[\w:-]+/.exec(xml)!;
  return `${m[0]} ${attr}="${n}"${xml.slice(m[0].length)}`;
}

export function exportOds(doc: SheetDoc): Uint8Array {
  const wb = doc.wb;
  const files: Zippable = {};
  const manifestExtra: string[] = [];
  const fonts = new Set<string>(['Calibri']);

  // number formats (content.xml) and cell styles
  const dataStyles = new Map<string, string>();
  let dataXml = '';
  const dataStyleFor = (fmt: string | undefined, prefix = 'N'): string | undefined => {
    if (!fmt || /^general$/i.test(fmt)) return undefined;
    const key = `${prefix}|${fmt}`;
    let name = dataStyles.get(key);
    if (name === undefined) {
      name = `${prefix}${dataStyles.size + 1}`;
      const xml = numFmtToOds(fmt, name);
      if (!xml) name = '';
      dataStyles.set(key, name);
      if (prefix === 'N') dataXml += xml;
      else cfDataXml += xml;
    }
    return name || undefined;
  };
  let cfDataXml = '';
  const cellStyles = new Map<number, string>();
  let cellXml = '';
  const cellStyleFor = (id: number | undefined): string => {
    if (!id) return 'Default';
    let name = cellStyles.get(id);
    if (!name) {
      name = `ce${cellStyles.size + 1}`;
      const st = wb.style(id);
      const ds = dataStyleFor(st.numFmt);
      cellXml += `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${ds ? ` style:data-style-name="${ds}"` : ''}>${cellStyleBody(st, fonts)}</style:style>`;
      cellStyles.set(id, name);
    }
    return name;
  };
  const colStyles = new Map<string, string>();
  let colXml = '';
  const colStyleFor = (w: number) => {
    const key = inch(w);
    let name = colStyles.get(key);
    if (!name) {
      name = `co${colStyles.size + 1}`;
      colXml += `<style:style style:name="${name}" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="${key}"/></style:style>`;
      colStyles.set(key, name);
    }
    return name;
  };
  const rowStyles = new Map<string, string>();
  let rowXml = '';
  const rowStyleFor = (h: number, optimal: boolean) => {
    const key = `${inch(h)}|${optimal}`;
    let name = rowStyles.get(key);
    if (!name) {
      name = `ro${rowStyles.size + 1}`;
      rowXml += `<style:style style:name="${name}" style:family="table-row"><style:table-row-properties style:row-height="${inch(h)}" fo:break-before="auto" style:use-optimal-row-height="${optimal}"/></style:style>`;
      rowStyles.set(key, name);
    }
    return name;
  };

  // conditional format styles (styles.xml)
  const cfStyles = new Map<string, string>();
  let cfStyleXml = '';
  const cfStyleFor = (cf: CondFormat): string | undefined => {
    if (cf.type === 'colorScale' || cf.type === 'dataBar' || cf.type === 'iconSet') return undefined;
    const st = cf.style ?? {};
    const key = JSON.stringify(st);
    let name = cfStyles.get(key);
    if (!name) {
      name = `AfficeCF${cfStyles.size + 1}`;
      const ds = dataStyleFor(st.numFmt, 'NC');
      cfStyleXml += `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${ds ? ` style:data-style-name="${ds}"` : ''}>${cellStyleBody(st, fonts)}</style:style>`;
      cfStyles.set(key, name);
    }
    return name;
  };

  // validations
  let validationsXml = '';
  const validationNames = new Map<Validation, string>();
  wb.sheets.forEach((sheet) => {
    for (const v of sheet.validations) {
      if (!v.ranges.length) continue;
      const name = `val${validationNames.size + 1}`;
      const xml = validationXml(v, name, sheet);
      if (!xml) continue;
      validationNames.set(v, name);
      validationsXml += xml;
    }
  });

  // page styles
  const pageLayouts = new Map<string, string>();
  let pageLayoutXmlAll = '';
  let masterXml = '';
  const masterFor = (sheet: Sheet): string => {
    const pr: SheetPrint = { ...DEFAULT_PRINT, ...(sheet.print ?? {}) };
    const { area: _area, repeatRows: _rows, ...page } = pr;
    const key = JSON.stringify(page);
    const known = pageLayouts.get(key);
    if (known) return known;
    const n = pageLayouts.size + 1;
    const master = n === 1 ? 'Default' : styleNameOf(`PageStyle_${sheet.name}`);
    pageLayoutXmlAll += pageLayoutXml(`pm${n}`, pr);
    masterXml += masterPageXml(master, n === 1 ? undefined : `PageStyle_${sheet.name}`, `pm${n}`, pr);
    pageLayouts.set(key, master);
    return master;
  };
  masterFor({ print: undefined } as Sheet);

  // pictures
  let pictureN = 0;
  const pictureFor = (src: string): { href: string; mime: string } | undefined => {
    const data = dataUrlToBytes(src);
    if (!data || !data.mime.startsWith('image/')) return undefined;
    const href = `Pictures/image${++pictureN}.${extFromMime(data.mime)}`;
    files[href] = [data.bytes, { level: 0 }];
    manifestExtra.push(`<manifest:file-entry manifest:full-path="${href}" manifest:media-type="${data.mime}"/>`);
    return { href, mime: data.mime };
  };

  let tableStyleXml = '';
  let graphicStyles = false;
  let objectN = 0;
  const tableParts: string[] = [];
  let dbRanges = '';
  let cellCount = 0;
  const now = new Date().toISOString().slice(0, 19);

  wb.sheets.forEach((sheet, sheetIndex) => {
    const ext = doc.engine.extent(sheet.id);
    const colCap = Math.min(MAX_COLS, Math.max(MIN_COLS, ext.cols + 1, ...[...sheet.cols.keys()].map((c) => c + 1)));
    const colDefault = (c: number) => sheet.cols.get(c)?.s ?? 0;

    // table style
    const ta = `ta${sheetIndex + 1}`;
    const tab = sheet.tabColor ? ` table:tab-color="${esc(sheet.tabColor)}" tableooo:tab-color="${esc(sheet.tabColor)}"` : '';
    tableStyleXml += `<style:style style:name="${ta}" style:family="table" style:master-page-name="${masterFor(sheet)}"><style:table-properties table:display="${!sheet.hidden}" style:writing-mode="${sheet.view.rtl ? 'rl-tb' : 'lr-tb'}"${tab}/></style:style>`;

    // merges and spilled values
    const mergeAt = new Map<number, Range>();
    const covered = new Set<number>();
    for (const m of sheet.merges) {
      if (m.r2 >= MAX_ROWS || m.c2 >= colCap) continue;
      mergeAt.set(cellKey(m.r1, m.c1), m);
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(cellKey(r, c));
    }
    const spillVals = new Map<number, Scalar>();
    const matrixAt = new Map<number, Range>();
    for (const [k, cell] of sheet.cells) {
      if (cell.f === undefined) continue;
      const r = keyRow(k);
      const c = keyCol(k);
      const area = doc.engine.spillAreaOf(sheet.id, r, c);
      if (!area || doc.engine.isSpillBlocked(sheet.id, r, c) || (area.r2 === area.r1 && area.c2 === area.c1)) continue;
      matrixAt.set(k, area);
      for (let rr = area.r1; rr <= area.r2; rr++)
        for (let cc = area.c1; cc <= area.c2; cc++) {
          const kk = cellKey(rr, cc);
          if (kk === k || sheet.cells.get(kk)) continue;
          const v = doc.value(sheet, rr, cc);
          if (v !== undefined && v !== null) spillVals.set(kk, v);
        }
    }

    // validations clipped to the written area
    const vals: { rg: Range; name: string }[] = [];
    for (const v of sheet.validations) {
      const name = validationNames.get(v);
      if (!name) continue;
      for (const rg of v.ranges) vals.push({ rg: { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, MAX_ROWS - 1), c2: Math.min(rg.c2, colCap - 1) }, name });
    }
    const validationAt = (r: number, c: number) => vals.find((v) => r >= v.rg.r1 && r <= v.rg.r2 && c >= v.rg.c1 && c <= v.rg.c2)?.name;

    // pictures anchored in cells
    const frames = new Map<number, string>();
    const colW = (c: number) => (sheet.isColHidden(c) ? 0 : sheet.colWidth(c));
    const rowH = (r: number) => (sheet.isRowHidden(r) ? 0 : sheet.rowHeight(r));
    const endOf = (start: number, offset: number, size: number, measure: (i: number) => number, max: number) => {
      let i = start;
      let pos = offset + size;
      while (i < max - 1 && pos > measure(i)) {
        pos -= measure(i);
        i++;
      }
      return { i, off: Math.max(0, pos) };
    };
    let z = 0;
    const frameXml = (anchor: ImageSpec['anchor'], w: number, h: number, style: string, name: string, inner: string) => {
      const endC = endOf(anchor.c, anchor.dx, w, colW, MAX_COLS);
      const endR = endOf(anchor.r, anchor.dy, h, rowH, MAX_ROWS);
      const k = cellKey(anchor.r, anchor.c);
      const xml = `<draw:frame table:end-cell-address="${esc(odsRangeAddress(sheet.name, endR.i, endC.i, endR.i, endC.i))}" table:end-x="${inch(endC.off)}" table:end-y="${inch(endR.off)}" draw:z-index="${z++}" draw:name="${esc(name)}" draw:style-name="${style}" draw:text-style-name="P1" svg:width="${inch(w)}" svg:height="${inch(h)}" svg:x="${inch(anchor.dx)}" svg:y="${inch(anchor.dy)}">${inner}</draw:frame>`;
      frames.set(k, (frames.get(k) ?? '') + xml);
      graphicStyles = true;
    };
    sheet.images.forEach((img: ImageSpec, n) => {
      const pic = pictureFor(img.src);
      if (!pic) return;
      frameXml(img.anchor, img.w, img.h, 'gr1', img.alt ? img.alt.slice(0, 60) : `Picture ${n + 1}`, `<draw:image xlink:href="${pic.href}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="${pic.mime}"><text:p/></draw:image>${img.alt ? `<svg:title>${esc(img.alt)}</svg:title>` : ''}`);
    });
    for (const spec of sheet.charts) {
      const obj = odsChartObject(doc, spec, sheet);
      if (!obj) continue;
      const name = `Object ${++objectN}`;
      files[`${name}/content.xml`] = strToU8(obj.content);
      files[`${name}/styles.xml`] = strToU8(obj.styles);
      manifestExtra.push(`<manifest:file-entry manifest:full-path="${name}/content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="${name}/styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="${name}/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.chart"/>`);
      frameXml(spec.anchor, spec.w, spec.h, 'gr2', spec.title?.slice(0, 60) || name, `<draw:object draw:notify-on-update-of-ranges="${esc(obj.ranges)}" xlink:href="./${name}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`);
    }

    // cells grouped by row
    const byRow = new Map<number, number[]>();
    const addCol = (r: number, c: number) => {
      let list = byRow.get(r);
      if (!list) byRow.set(r, (list = []));
      list.push(c);
    };
    for (const k of sheet.cells.keys()) if (keyCol(k) < colCap) addCol(keyRow(k), keyCol(k));
    for (const k of covered) addCol(keyRow(k), keyCol(k));
    for (const k of mergeAt.keys()) addCol(keyRow(k), keyCol(k));
    for (const k of spillVals.keys()) addCol(keyRow(k), keyCol(k));
    for (const k of frames.keys()) addCol(keyRow(k), keyCol(k));

    const rowBreaks = new Set<number>(byRow.keys());
    for (const r of sheet.rows.keys()) rowBreaks.add(r);
    for (const r of sheet.filterHidden) rowBreaks.add(r);
    for (const v of vals) {
      rowBreaks.add(v.rg.r1);
      if (v.rg.r2 + 1 < MAX_ROWS) rowBreaks.add(v.rg.r2 + 1);
    }
    const repeatRows = sheet.print?.repeatRows ?? 0;
    if (repeatRows > 0) rowBreaks.add(repeatRows);
    const sortedBreaks = [...rowBreaks].filter((r) => r >= 0 && r < MAX_ROWS).sort((a, b) => a - b);

    const cellXmlAt = (r: number, c: number): string => {
      const k = cellKey(r, c);
      const cell = sheet.cells.get(k);
      const isCovered = covered.has(k);
      const tag = isCovered ? 'table:covered-table-cell' : 'table:table-cell';
      const attrs: string[] = [];
      const sid = wb.cellStyleId(sheet, r, c);
      if (sid !== colDefault(c)) attrs.push(`table:style-name="${cellStyleFor(sid)}"`);
      const val = validationAt(r, c);
      if (val) attrs.push(`table:content-validation-name="${val}"`);
      const m = mergeAt.get(k);
      if (m) attrs.push(`table:number-columns-spanned="${m.c2 - m.c1 + 1}" table:number-rows-spanned="${m.r2 - m.r1 + 1}"`);
      const mx = matrixAt.get(k);
      if (mx) attrs.push(`table:number-matrix-columns-spanned="${mx.c2 - mx.c1 + 1}" table:number-matrix-rows-spanned="${mx.r2 - mx.r1 + 1}"`);
      let body = '';
      if (cell?.note) body += `<office:annotation office:display="false"><dc:date>${now}</dc:date>${paragraphs(cell.note)}</office:annotation>`;
      body += frames.get(k) ?? '';
      const formula = cell?.f !== undefined;
      const value = cell ? (formula ? doc.value(sheet, r, c) : cell.v) : spillVals.get(k);
      if (formula) attrs.push(`table:formula="${esc(toOdsFormula(cell!.f!))}"`);
      if (value !== undefined && value !== null) {
        cellCount++;
        const fmt = wb.style(sid).numFmt;
        if (isErr(value)) {
          const code = asErr(value).error === '#CIRC!' ? '#VALUE!' : asErr(value).error;
          if (!formula) attrs.push(`table:formula="${esc(code === '#N/A' ? 'of:=NA()' : `of:=${code}`)}"`);
          attrs.push('office:value-type="float" office:value="0" calcext:value-type="error"');
          body += paragraphs(code);
        } else if (typeof value === 'number') {
          const type = odsValueType(fmt);
          const date = type === 'date' ? serialToOdsDate(value) : undefined;
          // (calcext:value-type is only needed for text and errors)
          if (date) attrs.push(`office:value-type="date" office:date-value="${date}"`);
          else if (type === 'time') attrs.push(`office:value-type="time" office:time-value="${serialToOdsTime(value)}"`);
          else attrs.push(`office:value-type="${type === 'date' ? 'float' : type}" office:value="${value}"`);
          body += paragraphs(formatValue(value, fmt).text, cell?.link);
        } else if (typeof value === 'boolean') {
          attrs.push(`office:value-type="boolean" office:boolean-value="${value}"`);
          body += paragraphs(value ? 'TRUE' : 'FALSE', cell?.link);
        } else {
          const s = String(value);
          attrs.push(`office:value-type="string"${formula ? ` office:string-value="${esc(s)}"` : ''} calcext:value-type="string"`);
          body += paragraphs(s, cell?.link);
        }
      } else if (cell?.link) body += paragraphs(cell.link, cell.link);
      const open = `<${tag}${attrs.length ? ' ' + attrs.join(' ') : ''}`;
      return body ? `${open}>${body}</${tag}>` : `${open}/>`;
    };

    const emptyXml = (r: number, c: number, rowStyle: number | undefined): string => {
      const attrs: string[] = [];
      if (rowStyle !== undefined && rowStyle !== colDefault(c)) attrs.push(`table:style-name="${cellStyleFor(rowStyle)}"`);
      const val = validationAt(r, c);
      if (val) attrs.push(`table:content-validation-name="${val}"`);
      return `<table:table-cell${attrs.length ? ' ' + attrs.join(' ') : ''}/>`;
    };

    const rowXmlAt = (r: number): string => {
      const info = sheet.rows.get(r);
      const optimal = !info?.custom;
      const style = rowStyleFor(info?.h ?? sheet.defaultRowHeight, optimal);
      const vis = info?.hidden ? ' table:visibility="collapse"' : sheet.filterHidden.has(r) ? ' table:visibility="filter"' : '';
      const cols = new Set<number>(byRow.get(r) ?? []);
      for (const v of vals)
        if (r >= v.rg.r1 && r <= v.rg.r2) {
          cols.add(v.rg.c1);
          if (v.rg.c2 + 1 < colCap) cols.add(v.rg.c2 + 1);
        }
      const sorted = [...cols].filter((c) => c < colCap).sort((a, b) => a - b);
      const frags: Frag[] = [];
      const rowStyle = info?.s || undefined;
      // columns with their own default style break runs of row-styled cells
      const colBreaks = rowStyle ? [...sheet.cols.entries()].filter(([, ci]) => ci.s).map(([c]) => c) : [];
      const gap = (from: number, to: number) => {
        if (to < from) return;
        const stops = [from, ...colBreaks.filter((c) => c > from && c <= to).flatMap((c) => [c, c + 1]).filter((c) => c <= to), to + 1];
        const uniq = [...new Set(stops)].sort((a, b) => a - b);
        for (let i = 0; i < uniq.length - 1; i++) pushFrag(frags, emptyXml(r, uniq[i], rowStyle), uniq[i + 1] - uniq[i]);
      };
      let c = 0;
      for (const ic of sorted) {
        gap(c, ic - 1);
        pushFrag(frags, cellXmlAt(r, ic), 1);
        c = ic + 1;
      }
      gap(c, colCap - 1);
      // trailing plain cells add nothing
      while (frags.length > 1 && frags[frags.length - 1].xml === '<table:table-cell/>') frags.pop();
      const cells = frags.map((f) => withRepeat(f.xml, 'table:number-columns-repeated', f.n)).join('');
      return `<table:table-row table:style-name="${style}"${vis}>${cells}</table:table-row>`;
    };

    // rows: rows with content one by one, the rows between them as repeated blocks
    const rowFrags: { start: number; xml: string; n: number }[] = [];
    const addRows = (start: number, n: number) => {
      if (n <= 0) return;
      const xml = rowXmlAt(start);
      const last = rowFrags[rowFrags.length - 1];
      const sameGroup = !repeatRows || (last && (last.start < repeatRows) === (start < repeatRows));
      if (last && last.xml === xml && sameGroup) last.n += n;
      else rowFrags.push({ start, xml, n });
    };
    let next = 0;
    for (const r of sortedBreaks) {
      if (r > next) addRows(next, r - next);
      addRows(r, 1);
      next = r + 1;
    }
    if (next < MAX_ROWS) addRows(next, MAX_ROWS - next);
    const rowParts: string[] = [];
    let inHeader = false;
    for (const f of rowFrags) {
      const header = f.start < repeatRows;
      if (header && !inHeader) rowParts.push('<table:table-header-rows>');
      if (!header && inHeader) rowParts.push('</table:table-header-rows>');
      inHeader = header;
      rowParts.push(withRepeat(f.xml, 'table:number-rows-repeated', f.n));
    }
    if (inHeader) rowParts.push('</table:table-header-rows>');

    // columns
    const colFrags: Frag[] = [];
    const colBreaks = [...sheet.cols.keys()].filter((c) => c < colCap).sort((a, b) => a - b);
    const colXmlAt = (c: number) => {
      const ci = sheet.cols.get(c);
      return `<table:table-column table:style-name="${colStyleFor(ci?.w ?? sheet.defaultColWidth)}"${ci?.hidden ? ' table:visibility="collapse"' : ''} table:default-cell-style-name="${cellStyleFor(ci?.s)}"/>`;
    };
    let cNext = 0;
    for (const c of colBreaks) {
      if (c > cNext) pushFrag(colFrags, colXmlAt(cNext), c - cNext);
      pushFrag(colFrags, colXmlAt(c), 1);
      cNext = c + 1;
    }
    if (cNext < colCap) pushFrag(colFrags, colXmlAt(cNext), colCap - cNext);
    const colsXml = colFrags.map((f) => withRepeat(f.xml, 'table:number-columns-repeated', f.n)).join('');

    // conditional formats
    const cfs = sheet.cf
      .map((cf) => cfXml(cf, sheet, cfStyleFor(cf)))
      .filter(Boolean)
      .join('');

    // sheet-scoped names
    const local = wb.names.filter((n) => n.sheet === sheet.id).map((n) => namedXml(n, sheet.name));

    const attrs = [`table:name="${esc(sheet.name)}"`, `table:style-name="${ta}"`];
    if (sheet.protection?.enabled) attrs.push('table:protected="true"');
    const area = sheet.print?.area;
    if (area) attrs.push(`table:print-ranges="${esc(odsRangeAddress(sheet.name, area.r1, area.c1, Math.min(area.r2, MAX_ROWS - 1), Math.min(area.c2, MAX_COLS - 1)))}"`);
    tableParts.push(`<table:table ${attrs.join(' ')}>${colsXml}`, ...rowParts, `${local.length ? `<table:named-expressions>${local.join('')}</table:named-expressions>` : ''}${cfs ? `<calcext:conditional-formats>${cfs}</calcext:conditional-formats>` : ''}</table:table>`);

    if (sheet.filter) {
      const f = sheet.filter.range;
      dbRanges += `<table:database-range table:name="__Anonymous_Sheet_DB__${sheetIndex}" table:target-range-address="${esc(odsRangeAddress(sheet.name, f.r1, f.c1, Math.min(f.r2, MAX_ROWS - 1), Math.min(f.c2, MAX_COLS - 1)))}" table:display-filter-buttons="true">${filterXml(sheet.filter)}</table:database-range>`;
    }
  });

  const globalNames = wb.names.filter((n) => n.sheet === undefined).map((n) => namedXml(n, wb.sheets[0]?.name ?? 'Sheet1'));
  const graphics = graphicStyles
    ? '<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-horizontal-align="center" draw:textarea-vertical-align="middle" draw:color-mode="standard" draw:luminance="0%" draw:contrast="0%" draw:gamma="100%" draw:red="0%" draw:green="0%" draw:blue="0%" fo:clip="rect(0in, 0in, 0in, 0in)" draw:image-opacity="100%" style:mirror="none"/></style:style><style:style style:name="gr2" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-horizontal-align="center" draw:textarea-vertical-align="middle" draw:ole-draw-aspect="1"/></style:style><style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>'
    : '';

  // the body is kept in pieces: one string for a big sheet would be slow to encode
  const content = [
    `${XML_DECL}<office:document-content ${XMLNS} office:version="1.3"><office:scripts/>${fontFaces(fonts)}<office:automatic-styles>${colXml}${rowXml}${tableStyleXml}${dataXml}${cellXml}${graphics}</office:automatic-styles><office:body><office:spreadsheet><table:calculation-settings table:case-sensitive="false" table:search-criteria-must-apply-to-whole-cell="true" table:use-wildcards="true" table:use-regular-expressions="false" table:automatic-find-labels="false"/>${validationsXml ? `<table:content-validations>${validationsXml}</table:content-validations>` : ''}`,
    ...tableParts,
    `${globalNames.length ? `<table:named-expressions>${globalNames.join('')}</table:named-expressions>` : ''}${dbRanges ? `<table:database-ranges>${dbRanges}</table:database-ranges>` : ''}</office:spreadsheet></office:body></office:document-content>`,
  ];

  const styles = `${XML_DECL}<office:document-styles ${XMLNS} office:version="1.3">${fontFaces(fonts)}<office:styles><style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="0.5in"/><style:text-properties style:font-name="Calibri" fo:font-size="11pt" style:font-size-asian="11pt" style:font-size-complex="11pt"/></style:default-style><number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style><style:style style:name="Default" style:family="table-cell"><style:table-cell-properties style:vertical-align="bottom"/><style:text-properties style:font-name="Calibri" fo:font-size="11pt" style:font-size-asian="11pt" style:font-size-complex="11pt"/></style:style>${cfDataXml}${cfStyleXml}</office:styles><office:automatic-styles>${pageLayoutXmlAll}</office:automatic-styles><office:master-styles>${masterXml}</office:master-styles></office:document-styles>`;

  const props = wb.props;
  const created = props.created ? new Date(props.created).toISOString().slice(0, 19) : now;
  const meta = `${XML_DECL}<office:document-meta ${XMLNS} office:version="1.3"><office:meta><meta:generator>Affice</meta:generator>${props.title ? `<dc:title>${esc(props.title)}</dc:title>` : ''}${props.author ? `<meta:initial-creator>${esc(props.author)}</meta:initial-creator><dc:creator>${esc(props.author)}</dc:creator>` : ''}<meta:creation-date>${created}</meta:creation-date><dc:date>${now}</dc:date>${props.company ? `<meta:user-defined meta:name="Company">${esc(props.company)}</meta:user-defined>` : ''}<meta:document-statistic meta:table-count="${wb.sheets.length}" meta:cell-count="${cellCount}"/></office:meta></office:document-meta>`;

  files.mimetype = [strToU8(MIME), { level: 0 }];
  files['content.xml'] = encodeParts(content);
  files['styles.xml'] = strToU8(styles);
  files['meta.xml'] = strToU8(meta);
  files['settings.xml'] = strToU8(settingsXml(wb));
  files['META-INF/manifest.xml'] = strToU8(
    `${XML_DECL}<manifest:manifest xmlns:manifest="${NS_MANIFEST}" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>${manifestExtra.join('')}</manifest:manifest>`,
  );
  // the mimetype entry must come first and be stored uncompressed
  const ordered: Zippable = { mimetype: files.mimetype };
  for (const [k, v] of Object.entries(files)) if (k !== 'mimetype') ordered[k] = v;
  return zipSync(ordered, { level: 6 });
}

function settingsXml(wb: Workbook): string {
  const item = (name: string, type: string, value: string | number | boolean) => `<config:config-item config:name="${name}" config:type="${type}">${typeof value === 'string' ? esc(value) : value}</config:config-item>`;
  const tables = wb.sheets
    .map((s) => {
      const v = s.view;
      const fr = Math.max(0, v.freezeRows | 0);
      const fc = Math.max(0, v.freezeCols | 0);
      const active = v.selection?.active ?? { r: 0, c: 0 };
      return `<config:config-item-map-entry config:name="${esc(s.name)}">${[
        item('CursorPositionX', 'int', active.c),
        item('CursorPositionY', 'int', active.r),
        item('HorizontalSplitMode', 'short', fc ? 2 : 0),
        item('VerticalSplitMode', 'short', fr ? 2 : 0),
        item('HorizontalSplitPosition', 'int', fc),
        item('VerticalSplitPosition', 'int', fr),
        item('ActiveSplitRange', 'short', fc ? 3 : 2),
        item('PositionLeft', 'int', fc ? 0 : v.scrollCol),
        item('PositionRight', 'int', fc ? fc + v.scrollCol : 0),
        item('PositionTop', 'int', 0),
        item('PositionBottom', 'int', fr + v.scrollRow),
        item('ZoomType', 'short', 0),
        item('ZoomValue', 'int', Math.round((v.zoom || 1) * 100)),
        item('PageViewZoomValue', 'int', 60),
        item('ShowGrid', 'boolean', v.showGrid !== false),
      ].join('')}</config:config-item-map-entry>`;
    })
    .join('');
  const active = wb.sheets[wb.activeSheet] ?? wb.sheets[0];
  const av = active?.view;
  return `${XML_DECL}<office:document-settings ${XMLNS} office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>${item('ViewId', 'string', 'view1')}<config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named>${item('ActiveTable', 'string', active?.name ?? '')}${item('ZoomType', 'short', 0)}${item('ZoomValue', 'int', Math.round((av?.zoom || 1) * 100))}${item('ShowZeroValues', 'boolean', true)}${item('ShowNotes', 'boolean', true)}${item('ShowGrid', 'boolean', av?.showGrid !== false)}${item('HasColumnRowHeaders', 'boolean', av?.showHeaders !== false)}${item('HasSheetTabs', 'boolean', true)}${item('IsOutlineSymbolsSet', 'boolean', true)}</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set><config:config-item-set config:name="ooo:configuration-settings">${item('ShowGrid', 'boolean', av?.showGrid !== false)}${item('HasColumnRowHeaders', 'boolean', av?.showHeaders !== false)}${item('HasSheetTabs', 'boolean', true)}${item('AutoCalculate', 'boolean', wb.calcAuto)}${item('ShowZeroValues', 'boolean', true)}</config:config-item-set></office:settings></office:document-settings>`;
}

/* ============================================================ reading: helpers */

function kids(el: Element | Document): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
  return out;
}

function child(el: Element, ns: string, name: string): Element | undefined {
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (n as Element).localName === name && (n as Element).namespaceURI === ns) return n as Element;
  return undefined;
}

function at(el: Element, ns: string, name: string): string | undefined {
  return el.hasAttributeNS(ns, name) ? el.getAttributeNS(ns, name) ?? undefined : undefined;
}

function attrs(el: Element | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  if (!el) return out;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    out[a.localName ?? a.name] = a.value;
  }
  return out;
}

function descendants(root: Document | Element, ns: string, name: string): Element[] {
  return Array.from(root.getElementsByTagNameNS(ns, name));
}

/**
 * The parts of a document that hold styles. Searching only these keeps the reader from walking every
 * cell (and, in flat files, the charts' own styles stay out).
 */
function styleSections(d: Document | null): Element[] {
  if (!d?.documentElement) return [];
  return kids(d.documentElement).filter((k) => k.namespaceURI === NS.office && /^(font-face-decls|styles|automatic-styles|master-styles)$/.test(k.localName));
}

function inSections(sections: Element[], ns: string, name: string): Element[] {
  return sections.flatMap((sec) => descendants(sec, ns, name));
}

/** The text of a paragraph-bearing element, with text:s, text:tab and line breaks expanded. */
function textOf(el: Element): { text: string; link?: string } {
  let link: string | undefined;
  const walk = (e: Element): string => {
    let s = '';
    for (let n = e.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3 || n.nodeType === 4) {
        s += (n.nodeValue ?? '').replace(/[\r\n\t]+/g, ' ');
        continue;
      }
      if (n.nodeType !== 1) continue;
      const k = n as Element;
      if (k.namespaceURI === NS.office && k.localName === 'annotation') continue;
      if (k.namespaceURI === NS.draw) continue;
      switch (k.localName) {
        case 's':
          s += ' '.repeat(Number(at(k, NS.text, 'c') ?? 1) || 1);
          break;
        case 'tab':
          s += '\t';
          break;
        case 'line-break':
          s += '\n';
          break;
        case 'a':
          link ??= at(k, NS.xlink, 'href');
          s += walk(k);
          break;
        default:
          s += walk(k);
      }
    }
    return s;
  };
  const lines: string[] = [];
  for (const p of kids(el)) if (p.namespaceURI === NS.text && (p.localName === 'p' || p.localName === 'h')) lines.push(walk(p));
  return { text: lines.join('\n'), link };
}

const intAttr = (el: Element, name: string) => Math.max(1, parseInt(at(el, NS.table, name) ?? '1', 10) || 1);

/** Everything the reader needs from a cell's attributes, in one pass (cells are the bulk of a file). */
function cellAttrs(el: Element): CellAttrs {
  const out: CellAttrs = { repeat: 1, spanC: 1, spanR: 1, mxC: 1, mxR: 1 };
  const list = el.attributes;
  const int = (v: string) => Math.max(1, parseInt(v, 10) || 1);
  for (let i = 0; i < list.length; i++) {
    const attr = list[i];
    const ns = attr.namespaceURI;
    const v = attr.value;
    if (ns === NS.table)
      switch (attr.localName) {
        case 'number-columns-repeated':
          out.repeat = int(v);
          break;
        case 'style-name':
          out.style = v;
          break;
        case 'content-validation-name':
          out.validation = v;
          break;
        case 'formula':
          out.formula = v;
          break;
        case 'number-columns-spanned':
          out.spanC = int(v);
          break;
        case 'number-rows-spanned':
          out.spanR = int(v);
          break;
        case 'number-matrix-columns-spanned':
          out.mxC = int(v);
          break;
        case 'number-matrix-rows-spanned':
          out.mxR = int(v);
          break;
        default:
          break;
      }
    else if (ns === NS.office)
      switch (attr.localName) {
        case 'value-type':
          out.type = v;
          break;
        case 'value':
          out.value = v;
          break;
        case 'date-value':
          out.dateValue = v;
          break;
        case 'time-value':
          out.timeValue = v;
          break;
        case 'boolean-value':
          out.boolValue = v;
          break;
        case 'string-value':
          out.stringValue = v;
          break;
        default:
          break;
      }
    else if (ns === NS.calcext && attr.localName === 'value-type') out.calcType = v;
  }
  return out;
}

/** A row read through the DOM, in the same form the fast reader produces. */
function rowFromDom(el: Element): RowData {
  const cells: CellData[] = [];
  for (let k = el.firstChild; k; k = k.nextSibling) {
    if (k.nodeType !== 1) continue;
    const e = k as Element;
    if (e.namespaceURI === NS.table && (e.localName === 'table-cell' || e.localName === 'covered-table-cell')) cells.push(cellFromDom(e));
  }
  return { repeat: intAttr(el, 'number-rows-repeated'), style: at(el, NS.table, 'style-name'), visibility: at(el, NS.table, 'visibility'), defaultStyle: at(el, NS.table, 'default-cell-style-name'), cells };
}

function cellFromDom(cellEl: Element): CellData {
  let paras = 0;
  let marked = false;
  let note: Element | undefined;
  const frames: Element[] = [];
  for (let k = cellEl.firstChild; k; k = k.nextSibling) {
    if (k.nodeType !== 1) continue;
    const e = k as Element;
    if (e.namespaceURI === NS.text && (e.localName === 'p' || e.localName === 'h')) {
      paras++;
      // plain paragraphs hold a single run of text
      if (e.firstChild && (e.firstChild.nodeType !== 3 || e.firstChild.nextSibling)) marked = true;
    } else if (e.namespaceURI === NS.office && e.localName === 'annotation') note = e;
    else if (e.namespaceURI === NS.draw && e.localName === 'frame') frames.push(e);
  }
  return { a: cellAttrs(cellEl), paras, marked, note: note ? textOf(note).text : undefined, frames: () => frames, text: () => textOf(cellEl) };
}

/** Splits at a separator outside quotes, brackets and parentheses. */
function splitTop(s: string, sep: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let q = false;
  let cur = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') q = !q;
    else if (!q && (ch === '(' || ch === '[' || ch === '{')) depth++;
    else if (!q && (ch === ')' || ch === ']' || ch === '}')) depth--;
    if (!q && depth === 0 && s.startsWith(sep, i)) {
      out.push(cur);
      cur = '';
      i += sep.length - 1;
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out.map((x) => x.trim());
}

/** An OpenFormula condition operand → the model's operand text (number, text or "=formula"). */
function operandIn(src: string): string {
  const f = fromOdsFormula(src.trim());
  if (/^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(f)) return f;
  if (/^"([^"]|"")*"$/.test(f)) return f.slice(1, -1).replace(/""/g, '"');
  return '=' + f;
}

/** Parses "Sheet1.A1:Sheet1.B5" (or "$'My sheet'.$A$1:.$B$5") into a sheet name and range. */
function parseRangeAddress(addr: string): { sheet?: string; range: Range } | undefined {
  const ex = odsRefToExcel(addr.trim());
  const m = /^(?:('(?:[^']|'')+'|[^!]+)!)?(.+)$/.exec(ex);
  if (!m) return undefined;
  const sheet = m[1] ? (m[1].startsWith("'") ? m[1].slice(1, -1).replace(/''/g, "'") : m[1]) : undefined;
  const parts = m[2].replace(/\$/g, '').split(':');
  const cell = (s: string) => {
    const cm = /^([A-Za-z]{1,3})(\d+)$/.exec(s);
    return cm ? { r: Number(cm[2]) - 1, c: colIndex(cm[1]) } : undefined;
  };
  const colOnly = (s: string) => (/^[A-Za-z]{1,3}$/.test(s) ? colIndex(s) : undefined);
  if (parts.length === 2 && colOnly(parts[0]) !== undefined && colOnly(parts[1]) !== undefined) return { sheet, range: { r1: 0, c1: colOnly(parts[0])!, r2: MAX_ROWS - 1, c2: colOnly(parts[1])! } };
  const a = cell(parts[0]);
  const b = parts[1] !== undefined ? cell(parts[1]) : a;
  if (!a || !b) return undefined;
  return { sheet, range: { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.min(MAX_ROWS - 1, Math.max(a.r, b.r)), c2: Math.min(MAX_COLS - 1, Math.max(a.c, b.c)) } };
}

/** Collects rectangles (per key) and merges neighbours into as few ranges as possible. */
class RangeSet {
  private map = new Map<string, Range[]>();
  add(key: string, rg: Range): void {
    let list = this.map.get(key);
    if (!list) this.map.set(key, (list = []));
    const last = list[list.length - 1];
    // extend the previous rectangle to the right or downwards when it lines up
    if (last && last.r1 === rg.r1 && last.r2 === rg.r2 && last.c2 + 1 === rg.c1) last.c2 = rg.c2;
    else list.push({ ...rg });
  }
  get(key: string): Range[] {
    const list = this.map.get(key) ?? [];
    // merge vertically adjacent rectangles with the same columns
    list.sort((a, b) => a.c1 - b.c1 || a.c2 - b.c2 || a.r1 - b.r1);
    const out: Range[] = [];
    for (const rg of list) {
      const last = out[out.length - 1];
      if (last && last.c1 === rg.c1 && last.c2 === rg.c2 && last.r2 + 1 >= rg.r1) last.r2 = Math.max(last.r2, rg.r2);
      else out.push({ ...rg });
    }
    return out.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
  }
  keys(): string[] {
    return [...this.map.keys()];
  }
}

/* ============================================================ reading: styles */

interface OdStyle {
  name: string;
  display?: string;
  parent?: string;
  cell: Record<string, string>;
  para: Record<string, string>;
  text: Record<string, string>;
  dataStyle?: string;
  maps: { cond: string; style: string; base?: string }[];
}

const BORDER_KIND: Record<string, 'dashed' | 'dotted' | 'dashDot' | 'dashDotDot' | 'double'> = {
  dashed: 'dashed',
  'fine-dashed': 'dashed',
  dotted: 'dotted',
  'dash-dot': 'dashDot',
  'dash-dot-dot': 'dashDotDot',
  double: 'double',
  'double-thin': 'double',
};

function borderIn(v: string | undefined): Border | undefined {
  if (!v) return undefined;
  const parts = v.trim().split(/\s+/);
  let widthPt = 0.74;
  let kind = 'solid';
  let color: string | undefined;
  for (const p of parts) {
    if (/^#[0-9a-f]{6}$/i.test(p)) color = p.toLowerCase();
    else if (/^[\d.]+[a-z]+$/i.test(p)) widthPt = ((lengthPx(p) ?? 1) * 72) / 96;
    else kind = p.toLowerCase();
  }
  if (kind === 'none' || kind === 'hidden' || widthPt === 0) return undefined;
  const weight = widthPt >= 2.2 ? 2 : widthPt >= 1.2 ? 1 : 0;
  let style: BorderStyle;
  const k = BORDER_KIND[kind];
  if (k === 'double') style = 'double';
  else if (k === 'dotted') style = 'dotted';
  else if (k === 'dashed') style = weight ? 'mediumDashed' : 'dashed';
  else if (k === 'dashDot') style = weight ? 'mediumDashDot' : 'dashDot';
  else if (k === 'dashDotDot') style = weight ? 'mediumDashDotDot' : 'dashDotDot';
  else style = weight === 2 ? 'thick' : weight === 1 ? 'medium' : widthPt < 0.5 ? 'hair' : 'thin';
  return color && color !== '#000000' ? { style, color } : { style };
}

function colorIn(v: string | undefined): string | undefined {
  if (!v || v === 'transparent' || !/^#[0-9a-f]{6}$/i.test(v)) return undefined;
  return v.toLowerCase();
}

class StyleReader {
  styles = new Map<string, OdStyle>();
  byDisplay = new Map<string, OdStyle>();
  dataStyles = new Map<string, Element>();
  fonts = new Map<string, string>();
  defaults: Record<string, string> = {};
  private resolved = new Map<string, OdStyle>();
  private numFmts = new Map<string, string | undefined>();

  collect(root: Document | null): void {
    if (!root) return;
    const sections = styleSections(root);
    for (const ff of inSections(sections, NS.style, 'font-face')) {
      const name = at(ff, NS.style, 'name');
      const fam = at(ff, NS.svg, 'font-family')?.replace(/^['"]|['"]$/g, '');
      if (name) this.fonts.set(name, fam || name);
    }
    for (const ds of inSections(sections, NS.style, 'default-style'))
      if (at(ds, NS.style, 'family') === 'table-cell') Object.assign(this.defaults, attrs(child(ds, NS.style, 'text-properties')));
    for (const s of inSections(sections, NS.style, 'style')) {
      if (at(s, NS.style, 'family') !== 'table-cell') continue;
      const name = at(s, NS.style, 'name');
      if (!name) continue;
      const st: OdStyle = {
        name,
        display: at(s, NS.style, 'display-name'),
        parent: at(s, NS.style, 'parent-style-name'),
        cell: attrs(child(s, NS.style, 'table-cell-properties')),
        para: attrs(child(s, NS.style, 'paragraph-properties')),
        text: attrs(child(s, NS.style, 'text-properties')),
        dataStyle: at(s, NS.style, 'data-style-name'),
        maps: kids(s)
          .filter((k) => k.localName === 'map' && k.namespaceURI === NS.style)
          .map((m) => ({ cond: at(m, NS.style, 'condition') ?? '', style: at(m, NS.style, 'apply-style-name') ?? '', base: at(m, NS.style, 'base-cell-address') })),
      };
      this.styles.set(name, st);
      this.byDisplay.set(st.display ?? name, st);
    }
    for (const el of sections) {
      if (el.localName !== 'styles' && el.localName !== 'automatic-styles') continue;
      for (const ds of kids(el)) if (ds.namespaceURI === NS_NUMBER && ds.localName.endsWith('-style')) {
        const name = at(ds, NS.style, 'name');
        if (name) this.dataStyles.set(name, ds);
      }
    }
  }

  /** A style merged with its parents. */
  resolve(name: string | undefined): OdStyle | undefined {
    if (!name) return undefined;
    const hit = this.resolved.get(name);
    if (hit) return hit;
    const own = this.styles.get(name);
    if (!own) return undefined;
    this.resolved.set(name, own); // guards against parent loops
    const parent = own.parent && own.parent !== name ? this.resolve(own.parent) : undefined;
    const out: OdStyle = parent
      ? { ...own, cell: { ...parent.cell, ...own.cell }, para: { ...parent.para, ...own.para }, text: { ...parent.text, ...own.text }, dataStyle: own.dataStyle ?? parent.dataStyle }
      : own;
    this.resolved.set(name, out);
    return out;
  }

  numFmt(dataStyle: string | undefined): string | undefined {
    if (!dataStyle) return undefined;
    if (this.numFmts.has(dataStyle)) return this.numFmts.get(dataStyle);
    const el = this.dataStyles.get(dataStyle);
    let fmt: string | undefined;
    try {
      fmt = el ? odsToNumFmt(el, (n) => this.dataStyles.get(n)) : undefined;
    } catch {
      fmt = undefined;
    }
    this.numFmts.set(dataStyle, fmt);
    return fmt;
  }

  /** Converts style properties to the model (only what differs from the workbook default). */
  toCellStyle(st: Pick<OdStyle, 'cell' | 'para' | 'text' | 'dataStyle'>, withDefaults: boolean): CellStyle {
    const out: CellStyle = {};
    const text = withDefaults ? { ...this.defaults, ...st.text } : st.text;
    const fontName = text['font-name'] ? this.fonts.get(text['font-name']) ?? text['font-name'] : text['font-family']?.replace(/^['"]|['"]$/g, '');
    if (fontName && fontName !== 'Calibri') out.font = fontName;
    const size = text['font-size'] && text['font-size'].endsWith('pt') ? parseFloat(text['font-size']) : undefined;
    if (size && size !== 11) out.size = size;
    if (/bold|^[6-9]00$/.test(text['font-weight'] ?? '')) out.bold = true;
    if (/italic|oblique/.test(text['font-style'] ?? '')) out.italic = true;
    const ul = text['text-underline-style'];
    if (ul && ul !== 'none') out.underline = text['text-underline-type'] === 'double' ? 'double' : 'single';
    const lt = text['text-line-through-style'];
    if (lt && lt !== 'none') out.strike = true;
    const color = text['use-window-font-color'] === 'true' ? undefined : colorIn(text.color);
    if (color && color !== '#000000') out.color = color;
    const c = st.cell;
    const fill = colorIn(c['background-color']);
    if (fill) out.fill = fill;
    const all = borderIn(c.border);
    const side = (s: string) => (c[`border-${s}`] !== undefined ? borderIn(c[`border-${s}`]) : all);
    const bt = side('top');
    const bb = side('bottom');
    const bl = side('left');
    const br = side('right');
    if (bt) out.bt = bt;
    if (bb) out.bb = bb;
    if (bl) out.bl = bl;
    if (br) out.br = br;
    if (c['wrap-option'] === 'wrap') out.wrap = true;
    if (c['shrink-to-fit'] === 'true') out.shrink = true;
    const va = c['vertical-align'];
    if (va === 'top' || va === 'middle') out.vAlign = va;
    if (c.direction === 'ttb') out.rotation = 255;
    else if (c['rotation-angle']) {
      let a = Math.round(parseFloat(c['rotation-angle'])) % 360;
      if (a > 180) a -= 360;
      a = Math.max(-90, Math.min(90, a));
      if (a) out.rotation = a;
    }
    const protect = c['cell-protect'];
    if (protect) {
      if (protect === 'none' || protect === 'formula-hidden') out.locked = false;
      if (protect.includes('formula-hidden') || protect === 'hidden-and-protected') out.hideFormula = true;
    }
    if (c['repeat-content'] === 'true') out.hAlign = 'fill';
    else if (c['text-align-source'] !== 'value-type') {
      const ta = st.para['text-align'];
      if (ta === 'start' || ta === 'left') out.hAlign = 'left';
      else if (ta === 'center') out.hAlign = 'center';
      else if (ta === 'end' || ta === 'right') out.hAlign = 'right';
      else if (ta === 'justify') out.hAlign = 'justify';
    }
    const ml = lengthPx(st.para['margin-left']);
    if (ml && ml > 1) out.indent = Math.max(1, Math.round(ml / INDENT_PX));
    const fmt = this.numFmt(st.dataStyle);
    if (fmt) out.numFmt = fmt;
    return out;
  }
}

/* ============================================================ reading: conditions */

const OP_IN: Record<string, CompareOp> = { '=': 'equal', '==': 'equal', '!=': 'notEqual', '<>': 'notEqual', '>': 'greater', '<': 'less', '>=': 'greaterEqual', '<=': 'lessEqual' };

function compareIn(s: string, fn: string): { op: CompareOp; values: string[] } | undefined {
  const between = new RegExp(`^${fn}-is-(not-)?between\\((.*)\\)$`).exec(s);
  if (between) {
    let args = splitTop(between[2], ',');
    if (args.length < 2) args = splitTop(between[2], ';');
    return { op: between[1] ? 'notBetween' : 'between', values: args.slice(0, 2).map(operandIn) };
  }
  const m = new RegExp(`^${fn}\\(\\)\\s*(<=|>=|!=|<>|==|=|<|>)(.*)$`).exec(s);
  if (m) return { op: OP_IN[m[1]], values: [operandIn(m[2])] };
  return undefined;
}

function validationIn(cond: string): Pick<Validation, 'type' | 'op' | 'values'> | undefined {
  const s = cond.replace(/^[a-z][\w-]*:/i, '').trim();
  const parts = splitTop(s, ' and ');
  const first = parts[0];
  const list = /^cell-content-is-in-list\((.*)\)$/.exec(first);
  if (list) {
    const items = splitTop(list[1], ';').filter((x) => x !== '');
    if (items.length === 2 && /^TRUE\(\)$/i.test(items[0]) && /^FALSE\(\)$/i.test(items[1])) return { type: 'checkbox', values: [] };
    if (items.length === 1 && !/^"/.test(items[0]) && !/^-?[\d.]+$/.test(items[0])) return { type: 'list', values: ['=' + fromOdsFormula(items[0])] };
    return { type: 'list', values: items.map((x) => operandIn(x).replace(/^=/, '')) };
  }
  const custom = /^is-true-formula\((.*)\)$/.exec(first);
  if (custom) return { type: 'custom', values: ['=' + fromOdsFormula(custom[1])] };
  const typed: Record<string, ValidationType> = {
    'cell-content-is-whole-number()': 'whole',
    'cell-content-is-decimal-number()': 'decimal',
    'cell-content-is-date()': 'date',
    'cell-content-is-time()': 'time',
  };
  if (typed[first]) {
    const cmp = parts[1] ? compareIn(parts[1], 'cell-content') : undefined;
    return { type: typed[first], op: cmp?.op ?? 'greaterEqual', values: cmp?.values ?? ['-1E+307'] };
  }
  const len = compareIn(first, 'cell-content-text-length');
  if (len) return { type: 'textLength', ...len };
  const plain = compareIn(first, 'cell-content');
  if (plain) return { type: 'decimal', ...plain };
  return undefined;
}

const DATE_PERIOD_IN: Record<string, NonNullable<CondFormat['datePeriod']>> = Object.fromEntries(Object.entries(DATE_PERIOD_OUT).map(([k, v]) => [v, k as NonNullable<CondFormat['datePeriod']>]));

/** A calcext:condition value ("&gt;100", "between(1,5)", "contains-text(\"x\")"…) → rule fields. */
function cfConditionIn(value: string): Partial<CondFormat> | undefined {
  const v = value.trim();
  const call = /^([a-z-]+)\((.*)\)$/s.exec(v);
  const unquote = (s: string) => {
    const t = s.trim();
    return /^".*"$/s.test(t) ? t.slice(1, -1).replace(/""/g, '"') : fromOdsFormula(t);
  };
  if (call) {
    const [, fn, arg] = call;
    switch (fn) {
      case 'between':
      case 'not-between': {
        let args = splitTop(arg, ',');
        if (args.length < 2) args = splitTop(arg, ';');
        return { type: 'cell', op: fn === 'between' ? 'between' : 'notBetween', values: args.slice(0, 2).map(operandIn) };
      }
      case 'contains-text':
        return { type: 'text', text: unquote(arg) };
      case 'not-contains-text':
        return { type: 'notText', text: unquote(arg) };
      case 'begins-with':
        return { type: 'begins', text: unquote(arg) };
      case 'ends-with':
        return { type: 'ends', text: unquote(arg) };
      case 'top-elements':
      case 'bottom-elements':
      case 'top-percent':
      case 'bottom-percent':
        return { type: 'top', rank: Math.max(1, Math.round(Number(fromOdsFormula(arg)) || 10)), ...(fn.startsWith('bottom') ? { bottom: true } : {}), ...(fn.endsWith('percent') ? { percent: true } : {}) };
      case 'formula-is': {
        const f = fromOdsFormula(arg);
        const blank = /^LEN\(TRIM\(\$?[A-Z]{1,3}\$?\d+\)\)(=0|>0)$/i.exec(f.replace(/\s/g, ''));
        if (blank) return { type: blank[1] === '=0' ? 'blank' : 'notBlank' };
        return { type: 'formula', values: ['=' + f] };
      }
      default:
        return undefined;
    }
  }
  switch (v) {
    case 'duplicate':
      return { type: 'duplicate' };
    case 'unique':
      return { type: 'unique' };
    case 'above-average':
    case 'below-average':
    case 'above-equal-average':
    case 'below-equal-average':
      return { type: 'average', above: v.startsWith('above'), ...(v.includes('equal') ? { equalAverage: true } : {}) };
    case 'is-error':
      return { type: 'error' };
    case 'is-no-error':
      return { type: 'notError' };
    default:
      break;
  }
  const cmp = /^(<=|>=|!=|<>|==|=|<|>)(.*)$/s.exec(v);
  if (cmp) return { type: 'cell', op: OP_IN[cmp[1]], values: [operandIn(cmp[2])] };
  return undefined;
}

/** A legacy style:map condition ("cell-content()>100", "is-true-formula(…)"). */
function styleMapConditionIn(cond: string): Partial<CondFormat> | undefined {
  const s = cond.replace(/^[a-z][\w-]*:/i, '').trim();
  const f = /^is-true-formula\((.*)\)$/s.exec(s);
  if (f) return { type: 'formula', values: ['=' + fromOdsFormula(f[1])] };
  const cmp = compareIn(s, 'cell-content');
  if (cmp) return { type: 'cell', ...cmp };
  return undefined;
}

function condStyleIn(st: OdStyle | undefined, reader: StyleReader): CondFormat['style'] {
  if (!st) return undefined;
  // only what the rule's style sets itself (its parent is normally the Default style)
  const chain: OdStyle[] = [];
  for (let s: OdStyle | undefined = st; s && s.name !== 'Default' && chain.length < 10; s = s.parent ? reader.styles.get(s.parent) : undefined) chain.unshift(s);
  const merged = { cell: {}, para: {}, text: {}, dataStyle: undefined as string | undefined };
  for (const s of chain) {
    Object.assign(merged.cell, s.cell);
    Object.assign(merged.para, s.para);
    Object.assign(merged.text, s.text);
    merged.dataStyle = s.dataStyle ?? merged.dataStyle;
  }
  const full = reader.toCellStyle(merged, false);
  const out: NonNullable<CondFormat['style']> = {};
  for (const k of ['bold', 'italic', 'underline', 'strike', 'color', 'fill', 'numFmt', 'bt', 'br', 'bb', 'bl'] as const) if (full[k] !== undefined) (out as Record<string, unknown>)[k] = full[k];
  return Object.keys(out).length ? out : undefined;
}

/* ============================================================ reading: workbook */

interface PendingValidation {
  rule: Omit<Validation, 'ranges' | 'id'>;
}

function pageSetupIn(master: Element | undefined, layouts: Map<string, Element>): SheetPrint | undefined {
  if (!master) return undefined;
  const layout = layouts.get(at(master, NS.style, 'page-layout-name') ?? '');
  const p = attrs(layout ? child(layout, NS.style, 'page-layout-properties') : undefined);
  const pr: SheetPrint = { ...DEFAULT_PRINT };
  const w = lengthPx(p['page-width']);
  const h = lengthPx(p['page-height']);
  if (w && h) {
    const [a, b] = [Math.min(w, h) / 96, Math.max(w, h) / 96];
    let best: SheetPrint['paper'] = pr.paper;
    let err = Infinity;
    for (const [k, [pw, ph]] of Object.entries(PAPER_IN) as [SheetPrint['paper'], [number, number]][]) {
      const e = Math.abs(pw - a) + Math.abs(ph - b);
      if (e < err) {
        err = e;
        best = k;
      }
    }
    if (err < 0.3) pr.paper = best;
    pr.orientation = p['print-orientation'] === 'landscape' || w > h ? 'landscape' : 'portrait';
  }
  const mt = lengthPx(p['margin-top']);
  const ml = lengthPx(p['margin-left']);
  if (mt !== undefined && ml !== undefined) {
    let best: SheetPrint['margins'] = 'normal';
    let err = Infinity;
    for (const [k, [t, , , l]] of Object.entries(MARGINS_IN) as [SheetPrint['margins'], number[]][]) {
      const e = Math.abs(t - mt / 96) + Math.abs(l - ml / 96);
      if (e < err) {
        err = e;
        best = k;
      }
    }
    pr.margins = best;
  }
  if (p['scale-to-pages'] === '1') pr.fit = 'page';
  else if ((p['scale-to-X'] ?? '') === '1' && (p['scale-to-Y'] ?? '0') === '0') pr.fit = 'width';
  else if (p['scale-to']) {
    pr.fit = 'none';
    pr.scale = Math.max(10, Math.min(400, parseFloat(p['scale-to']) || 100));
  } else pr.fit = 'none';
  const print = p.print ?? '';
  pr.gridlines = /\bgrid\b/.test(print);
  pr.headings = /\bheaders\b/.test(print);
  if (p['table-centering'] === 'horizontal' || p['table-centering'] === 'both') pr.centerH = true;
  const hf = (tag: string): string | undefined => {
    const el = child(master, NS.style, tag);
    if (!el || at(el, NS.style, 'display') === 'false') return undefined;
    const region = child(el, NS.style, 'region-center') ?? child(el, NS.style, 'region-left') ?? child(el, NS.style, 'region-right') ?? el;
    let out = '';
    const walk = (e: Element) => {
      for (let n = e.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 3) out += n.nodeValue ?? '';
        else if (n.nodeType === 1) {
          const k = n as Element;
          const codes: Record<string, string> = { 'page-number': '&P', 'page-count': '&N', date: '&D', time: '&T', 'sheet-name': '&A', title: '&F', 'file-name': '&F' };
          if (codes[k.localName]) out += codes[k.localName];
          else if (k.localName === 's') out += ' '.repeat(Number(at(k, NS.text, 'c') ?? 1) || 1);
          else if (k.localName === 'p' && out) {
            out += ' ';
            walk(k);
          } else walk(k);
        }
      }
    };
    walk(region);
    return out.trim() || undefined;
  };
  pr.header = hf('header');
  pr.footer = hf('footer');
  const clean = (o: object) => JSON.stringify(Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined && v !== false).sort()));
  return clean(pr) === clean(DEFAULT_PRINT) ? undefined : pr;
}

export async function importOds(bytes: Uint8Array, options: { fast?: boolean } = {}): Promise<{ wb: Workbook; warning?: string }> {
  if (options.fast === false) return readOds(bytes, false);
  try {
    return readOds(bytes, true);
  } catch (e) {
    // the fast reader met something unusual: read the whole file through the DOM
    if (isFallback(e)) return readOds(bytes, false);
    throw e;
  }
}

function readOds(bytes: Uint8Array, fast: boolean): { wb: Workbook; warning?: string } {
  let files: Record<string, Uint8Array> = {};
  const parseXml = (xml: string): Document | null => {
    const d = new DOMParser().parseFromString(xml, 'application/xml');
    return d.getElementsByTagName('parsererror').length ? null : d;
  };
  let parse: (path: string) => Document | null;
  const zipped = bytes[0] === 0x50 && bytes[1] === 0x4b;
  // the text of content.xml and, on the fast path, where its rows are
  let contentText: string | undefined;
  let split: SplitContent | null = null;
  let content: Document | null;
  if (zipped) {
    try {
      files = unzipSync(bytes);
    } catch {
      throw new Error('This file is not a valid OpenDocument spreadsheet.');
    }
    parse = (path) => (files[path] ? parseXml(strFromU8(files[path])) : null);
    if (fast && files['content.xml']) {
      contentText = strFromU8(files['content.xml']);
      split = splitContent(contentText);
    }
    content = split ? parseXml(split.reduced) : parse('content.xml');
  } else {
    // flat OpenDocument (.fods): one XML file holds everything
    const text = strFromU8(bytes).replace(/^\uFEFF/, '');
    if (fast) {
      contentText = text;
      split = splitContent(text, true);
    }
    const flat = parseXml(split ? split.reduced : text);
    parse = () => flat;
    content = flat;
  }
  if (!content) {
    if (split) throw fallback();
    throw new Error('This file is not a valid OpenDocument spreadsheet.');
  }
  const stylesDoc = parse('styles.xml');
  const reader = new StyleReader();
  reader.collect(stylesDoc);
  reader.collect(content);
  const layouts = new Map<string, Element>();
  const masters = new Map<string, Element>();
  if (stylesDoc) {
    const sections = styleSections(stylesDoc);
    for (const l of inSections(sections, NS.style, 'page-layout')) layouts.set(at(l, NS.style, 'name') ?? '', l);
    for (const m of inSections(sections, NS.style, 'master-page')) masters.set(at(m, NS.style, 'name') ?? '', m);
  }
  const tableStyles = new Map<string, { props: Record<string, string>; master?: string }>();
  const contentStyles = inSections(styleSections(content), NS.style, 'style');
  for (const s of contentStyles)
    if (at(s, NS.style, 'family') === 'table') tableStyles.set(at(s, NS.style, 'name') ?? '', { props: attrs(child(s, NS.style, 'table-properties')), master: at(s, NS.style, 'master-page-name') });
  const colWidths = new Map<string, number>();
  const rowHeights = new Map<string, { h?: number; optimal: boolean }>();
  for (const s of contentStyles) {
    const family = at(s, NS.style, 'family');
    const name = at(s, NS.style, 'name') ?? '';
    if (family === 'table-column') {
      const w = lengthPx(at(child(s, NS.style, 'table-column-properties') ?? s, NS.style, 'column-width'));
      if (w !== undefined) colWidths.set(name, w);
    } else if (family === 'table-row') {
      const p = child(s, NS.style, 'table-row-properties');
      rowHeights.set(name, { h: p ? lengthPx(at(p, NS.style, 'row-height') ?? at(p, NS.style, 'min-row-height')) : undefined, optimal: p ? at(p, NS.style, 'use-optimal-row-height') !== 'false' : true });
    }
  }

  const wb = new Workbook();
  const warnings: string[] = [];
  let seqN = 0;
  const seq = () => `o${(seqN++).toString(36)}`;
  let cellTotal = 0;

  // cell styles → model style ids
  const defaultStyle = reader.toCellStyle({ cell: {}, para: {}, text: {}, ...(reader.resolve('Default') ?? {}) }, true);
  const defaultId = Object.keys(defaultStyle).length ? wb.styleId(defaultStyle) : 0;
  const styleIds = new Map<string, number>();
  const styleIdOf = (name: string | undefined): number => {
    if (!name || name === 'Default') return defaultId;
    let id = styleIds.get(name);
    if (id === undefined) {
      const st = reader.resolve(name);
      id = st ? wb.styleId(reader.toCellStyle(st, true)) : defaultId;
      styleIds.set(name, id);
    }
    return id;
  };

  // validations
  const validations = new Map<string, PendingValidation>();
  const spreadsheet = child(child(content.documentElement, NS.office, 'body') ?? content.documentElement, NS.office, 'spreadsheet');
  if (!spreadsheet) throw new Error('This file is not a valid OpenDocument spreadsheet.');
  for (const v of kids(child(spreadsheet, NS.table, 'content-validations') ?? spreadsheet).filter((k) => k.localName === 'content-validation')) {
    const name = at(v, NS.table, 'name');
    if (!name) continue;
    const cond = at(v, NS.table, 'condition');
    const parsed = cond ? validationIn(cond) : { type: 'any' as const, values: [] };
    if (!parsed) continue;
    const help = child(v, NS.table, 'help-message');
    const err = child(v, NS.table, 'error-message');
    const rule: PendingValidation['rule'] = {
      ...parsed,
      allowBlank: at(v, NS.table, 'allow-empty-cell') !== 'false',
      showDropdown: at(v, NS.table, 'display-list') !== 'none',
    };
    if (help && at(help, NS.table, 'display') !== 'false') {
      rule.promptTitle = at(help, NS.table, 'title') || undefined;
      rule.prompt = textOf(help).text || undefined;
    }
    if (err) {
      const t = at(err, NS.table, 'message-type');
      rule.errorStyle = t === 'warning' ? 'warning' : t === 'information' ? 'information' : 'stop';
      rule.errorTitle = at(err, NS.table, 'title') || undefined;
      rule.error = textOf(err).text || undefined;
    }
    validations.set(name, { rule });
  }

  const pictures = new Map<string, string>();
  const pictureSrc = (href: string | undefined): string | undefined => {
    if (!href) return undefined;
    const path = href.replace(/^\.\//, '');
    if (pictures.has(path)) return pictures.get(path);
    const data = files[path];
    const ext = path.split('.').pop()?.toLowerCase() ?? '';
    const mime = mimeFromExt(ext);
    const src = data && mime.startsWith('image/') && !['image/emf', 'image/wmf'].includes(mime) ? bytesToDataUrl(data, mime) : undefined;
    pictures.set(path, src ?? '');
    return src;
  };

  const tables = kids(spreadsheet).filter((t) => t.localName === 'table' && t.namespaceURI === NS.table);
  if (split && split.regions.length !== tables.length) throw fallback();
  const sheetByName = new Map<string, Sheet>();
  for (const [index, table] of tables.entries()) {
    const name = at(table, NS.table, 'name') || `Sheet${index + 1}`;
    const sheet = wb.addSheet(wb.uniqueSheetName(name.slice(0, 100)));
    sheetByName.set(name, sheet);
    const ts = tableStyles.get(at(table, NS.table, 'style-name') ?? '');
    if (ts) {
      if (ts.props.display === 'false') sheet.hidden = true;
      if (ts.props['writing-mode'] === 'rl-tb') sheet.view.rtl = true;
      const tab = colorIn(ts.props['tab-color']);
      if (tab) sheet.tabColor = tab;
      const pr = pageSetupIn(masters.get(ts.master ?? 'Default') ?? masters.get('Default'), layouts);
      if (pr) sheet.print = pr;
    }
    if (at(table, NS.table, 'protected') === 'true') sheet.protection = { enabled: true };
    const printRanges = at(table, NS.table, 'print-ranges');
    if (printRanges) {
      const pr = parseRangeAddress(splitRangeList(printRanges)[0] ?? '');
      if (pr) sheet.print = { ...(sheet.print ?? DEFAULT_PRINT), area: pr.range };
    }

    // columns
    const colRuns: { c: number; n: number; w?: number; hidden: boolean; style?: string }[] = [];
    let col = 0;
    const addColumn = (cd: ColData) => {
      if (col < MAX_COLS) colRuns.push({ c: col, n: Math.min(cd.repeat, MAX_COLS - col), w: colWidths.get(cd.style ?? ''), hidden: cd.visibility === 'collapse', style: cd.defaultStyle });
      col += cd.repeat;
    };
    const readCols = (el: Element) => {
      for (const k of kids(el)) {
        if (k.namespaceURI !== NS.table) continue;
        if (k.localName === 'table-column') addColumn({ repeat: intAttr(k, 'number-columns-repeated'), style: at(k, NS.table, 'style-name'), visibility: at(k, NS.table, 'visibility'), defaultStyle: at(k, NS.table, 'default-cell-style-name') });
        else if (k.localName === 'table-column-group' || k.localName === 'table-header-columns' || k.localName === 'table-columns') readCols(k);
      }
    };
    const colStyleId: number[] = [];
    let colsDone = false;
    // columns come before rows: their sizes and styles are settled when the first row arrives
    const finishColumns = () => {
      if (colsDone) return;
      colsDone = true;
      const widthCount = new Map<number, number>();
      for (const run of colRuns) if (run.w !== undefined) widthCount.set(Math.round(run.w), (widthCount.get(Math.round(run.w)) ?? 0) + run.n);
      const defaultW = [...widthCount.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
      if (defaultW) sheet.defaultColWidth = defaultW;
      for (const run of colRuns) {
        const sid = run.style && run.style !== 'Default' ? styleIdOf(run.style) : 0;
        const customW = run.w !== undefined && Math.abs(run.w - sheet.defaultColWidth) > 0.5;
        if (!customW && !run.hidden && !sid) continue;
        if (run.n > 2000 && !run.hidden && !sid) continue;
        for (let c = run.c; c < run.c + Math.min(run.n, 16384); c++) {
          const info: { w?: number; hidden?: boolean; s?: number } = {};
          if (customW) info.w = Math.round(run.w!);
          if (run.hidden) info.hidden = true;
          if (sid && sid !== defaultId) {
            info.s = sid;
            colStyleId[c] = sid;
          }
          if (Object.keys(info).length) sheet.cols.set(c, info);
        }
      }
    };
    const colDefaultName = (c: number): string | undefined => {
      let lo = 0;
      let hi = colRuns.length - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const run = colRuns[mid];
        if (c < run.c) hi = mid - 1;
        else if (c >= run.c + run.n) lo = mid + 1;
        else return run.style;
      }
      return undefined;
    };

    // rows and cells
    const rowRuns: { r: number; n: number; h?: number; optimal: boolean; hidden: boolean; filtered: boolean }[] = [];
    const cellRanges = new RangeSet();
    const spillAreas: { anchor: number; area: Range }[] = [];
    const legacyCf = new RangeSet();
    let headerRows: { start: number; end: number } | undefined;
    let row = 0;
    let capped = false;

    const readRow = (rd: RowData) => {
      finishColumns();
      const n = rd.repeat;
      const rs = rowHeights.get(rd.style ?? '');
      const vis = rd.visibility;
      if (row < MAX_ROWS) rowRuns.push({ r: row, n: Math.min(n, MAX_ROWS - row), h: rs?.h, optimal: rs?.optimal ?? true, hidden: vis === 'collapse', filtered: vis === 'filter' });
      const rowDefault = rd.defaultStyle;
      let c = 0;
      const cells = rd.cells;
      for (const [ci, cd] of cells.entries()) {
        const a = cd.a;
        const cn = a.repeat;
        const c0 = c;
        c += cn;
        if (c0 >= MAX_COLS || row >= MAX_ROWS) continue;
        const reps = Math.min(cn, MAX_COLS - c0);
        const styleName = a.style ?? rowDefault ?? colDefaultName(c0);
        const valName = a.validation;
        const formulaSrc = a.formula;
        const type = a.type;
        const calcType = a.calcType;
        const { paras, marked } = cd;
        const note = cd.note;
        const framesEls = cd.frames();
        const hasText = paras > 0;
        const { spanC, spanR, mxC, mxR } = a;
        const empty = !formulaSrc && !type && !hasText && note === undefined && !framesEls.length;
        const rowsHere = Math.min(n, MAX_ROWS - row);
        if (valName && validations.has(valName)) cellRanges.add(valName, { r1: row, c1: c0, r2: row + rowsHere - 1, c2: c0 + reps - 1 });
        if (empty) {
          if (!styleName || styleName === colDefaultName(c0) || styleName === 'Default') continue;
          const sid = styleIdOf(styleName);
          if (!sid || sid === defaultId) continue;
          // a styled run to the end of the row is row formatting
          if (ci === cells.length - 1 && reps > 256) {
            for (let r = row; r < row + Math.min(rowsHere, 100_000); r++) sheet.rows.set(r, { ...(sheet.rows.get(r) ?? {}), s: sid });
            continue;
          }
          if (reps * rowsHere > 100_000) continue;
          for (let r = row; r < row + rowsHere; r++) for (let cc = c0; cc < c0 + reps; cc++) if ((colStyleId[cc] ?? 0) !== sid) sheet.put(cellKey(r, cc), { s: sid });
          const st = reader.styles.get(styleName);
          if (st?.maps.length) legacyCf.add(styleName, { r1: row, c1: c0, r2: row + rowsHere - 1, c2: c0 + reps - 1 });
          continue;
        }
        // content
        const cell: { v?: Scalar; f?: string; s?: number; link?: string; note?: string } = {};
        const sid = styleIdOf(styleName);
        // numbers carry their value in attributes: their text is only needed for links or odd files
        const needText = hasText && (marked || paras > 1 || type === 'string' || type === undefined || calcType === 'error' || (type !== 'boolean' && a.value === undefined && a.dateValue === undefined && a.timeValue === undefined));
        const t = needText ? cd.text() : { text: '', link: undefined };
        if (t.link) cell.link = t.link;
        if (note !== undefined) cell.note = note;
        let extraFmt: string | undefined;
        if (calcType === 'error') cell.v = errFromCode((t.text.trim() || '#VALUE!').replace(/^Err:\d+$/, '#VALUE!'));
        else
          switch (type) {
            case 'float':
            case 'percentage':
            case 'currency': {
              const num = Number(a.value);
              cell.v = Number.isFinite(num) ? num : t.text;
              if (type === 'percentage') extraFmt = '0%';
              break;
            }
            case 'date': {
              const serial = odsDateToSerial(a.dateValue ?? '');
              cell.v = serial ?? t.text;
              if (serial !== undefined) extraFmt = serial % 1 ? `${SHORT_DATE} h:mm` : SHORT_DATE;
              break;
            }
            case 'time': {
              const serial = odsTimeToSerial(a.timeValue ?? '');
              cell.v = serial ?? t.text;
              if (serial !== undefined) extraFmt = serial >= 1 ? '[h]:mm:ss' : 'h:mm:ss';
              break;
            }
            case 'boolean': {
              const b = (a.boolValue ?? '').toLowerCase();
              cell.v = b === 'true' || b === '1';
              break;
            }
            case 'string':
              cell.v = a.stringValue ?? t.text;
              break;
            case 'void':
              break;
            default:
              if (t.text) cell.v = t.text;
          }
        if (formulaSrc) {
          const f = fromOdsFormula(formulaSrc);
          // TRUE() and FALSE() are how LibreOffice stores logical constants
          if (/^(TRUE|FALSE)$/i.test(f.trim())) cell.v = f.trim().toUpperCase() === 'TRUE';
          else if (f.trim()) cell.f = f;
          if (cell.f !== undefined && calcType !== 'error' && type === undefined && cell.v === '') delete cell.v;
        }
        let finalSid = sid;
        if (extraFmt && typeof cell.v === 'number' && !wb.style(sid).numFmt) finalSid = wb.patchStyle(sid, { numFmt: extraFmt });
        if (finalSid && finalSid !== (colStyleId[c0] ?? 0)) cell.s = finalSid;
        else if (!finalSid && colStyleId[c0]) cell.s = 0;
        if (cell.v === undefined) delete cell.v;
        const st = styleName ? reader.styles.get(styleName) : undefined;
        if (st?.maps.length) legacyCf.add(styleName!, { r1: row, c1: c0, r2: row + rowsHere - 1, c2: c0 + reps - 1 });
        for (let r = row; r < row + rowsHere; r++) {
          for (let cc = c0; cc < c0 + reps; cc++) {
            if (++cellTotal > MAX_CELLS) {
              capped = true;
              break;
            }
            const k = cellKey(r, cc);
            if (cell.v !== undefined || cell.f !== undefined || cell.s !== undefined || cell.link || cell.note) sheet.put(k, { ...cell });
            if (spanC > 1 || spanR > 1) sheet.merges.push({ r1: r, c1: cc, r2: Math.min(MAX_ROWS - 1, r + spanR - 1), c2: Math.min(MAX_COLS - 1, cc + spanC - 1) });
            if (cell.f !== undefined && (mxC > 1 || mxR > 1)) spillAreas.push({ anchor: k, area: { r1: r, c1: cc, r2: r + mxR - 1, c2: cc + mxC - 1 } });
            for (const fr of framesEls) readFrame(fr, sheet, r, cc);
          }
          if (capped) break;
        }
      }
      row += n;
    };

    const readFrame = (fr: Element, target: Sheet, r: number, c: number, x = lengthPx(at(fr, NS.svg, 'x')) ?? 0, y = lengthPx(at(fr, NS.svg, 'y')) ?? 0) => {
      // offsets are stored in inches: round away the conversion noise
      const dx = Math.max(0, Math.round(x * 10) / 10);
      const dy = Math.max(0, Math.round(y * 10) / 10);
      const obj = child(fr, NS.draw, 'object');
      if (obj) {
        // charts: embedded chart documents (inline in flat files)
        const inline = child(obj, NS.office, 'document');
        const href = at(obj, NS.xlink, 'href')?.replace(/^\.\//, '').replace(/\/$/, '');
        const root = inline ?? (zipped && href ? parse(`${href}/content.xml`) : null);
        const spec = root ? readOdsChart(root, target.name) : null;
        if (!spec) return;
        const w = lengthPx(at(fr, NS.svg, 'width')) ?? 480;
        const h = lengthPx(at(fr, NS.svg, 'height')) ?? 290;
        target.charts.push({ ...spec, id: seq(), anchor: { r, c, dx, dy }, w: Math.max(60, Math.round(w)), h: Math.max(40, Math.round(h)) });
        return;
      }
      const img = child(fr, NS.draw, 'image');
      if (!img) return;
      const binary = child(img, NS.office, 'binary-data')?.textContent?.replace(/\s+/g, '');
      const src = binary ? embeddedPicture(binary) : pictureSrc(at(img, NS.xlink, 'href'));
      if (!src) return;
      const w = lengthPx(at(fr, NS.svg, 'width')) ?? 200;
      const h = lengthPx(at(fr, NS.svg, 'height')) ?? 150;
      const alt = child(fr, NS.svg, 'title')?.textContent || child(fr, NS.svg, 'desc')?.textContent || undefined;
      target.images.push({ id: seq(), src, anchor: { r, c, dx, dy }, w: Math.max(10, Math.round(w)), h: Math.max(10, Math.round(h)), alt });
    };

    const shapes: Element[] = [];
    const localNames: Element[] = [];
    const cfBlocks: Element[] = [];
    let headerStart = 0;
    const header = (start: boolean) => {
      if (start) headerStart = row;
      else headerRows ??= { start: headerStart, end: row };
    };
    const readRows = (el: Element) => {
      for (const k of kids(el)) {
        if (k.namespaceURI === NS.table) {
          if (k.localName === 'table-row') readRow(rowFromDom(k));
          else if (k.localName === 'table-header-rows') {
            header(true);
            readRows(k);
            header(false);
          } else if (k.localName === 'table-row-group' || k.localName === 'table-rows') readRows(k);
          else if (k.localName === 'shapes') shapes.push(k);
          else if (k.localName === 'named-expressions') localNames.push(k);
        } else if (k.namespaceURI === NS.calcext && k.localName === 'conditional-formats') cfBlocks.push(k);
      }
    };
    readCols(table);
    readRows(table);
    // the fast path: this sheet's rows and columns were cut out of the DOM and are streamed instead
    const region = split?.regions[index];
    if (split && region)
      scanRegion(contentText!, region, split.prefixes, split.xmlns, {
        column: addColumn,
        row: readRow,
        headerRows: header,
        extra: (el) => {
          if (el.namespaceURI === NS.table && el.localName === 'shapes') shapes.push(el);
          else if (el.namespaceURI === NS.table && el.localName === 'named-expressions') localNames.push(el);
          else if (el.namespaceURI === NS.calcext && el.localName === 'conditional-formats') cfBlocks.push(el);
        },
      });
    finishColumns();
    if (capped) warnings.push(`“${sheet.name}” is very large; only the first 3 million cells were loaded.`);

    // row heights: the most common height becomes the default
    const heightCount = new Map<number, number>();
    for (const run of rowRuns) if (run.h !== undefined) heightCount.set(Math.round(run.h * 2) / 2, (heightCount.get(Math.round(run.h * 2) / 2) ?? 0) + run.n);
    const defaultH = [...heightCount.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
    if (defaultH) sheet.defaultRowHeight = Math.round(defaultH);
    for (const run of rowRuns) {
      const custom = run.h !== undefined && Math.abs(run.h - sheet.defaultRowHeight) > 0.5;
      if (!custom && !run.hidden && !run.filtered) continue;
      if (run.n > 100_000) continue;
      for (let r = run.r; r < run.r + run.n; r++) {
        const info = { ...(sheet.rows.get(r) ?? {}) };
        if (custom) {
          info.h = Math.round(run.h!);
          if (!run.optimal) info.custom = true;
        }
        if (run.hidden) info.hidden = true;
        if (run.filtered) sheet.filterHidden.add(r);
        if (Object.keys(info).length) sheet.rows.set(r, info);
      }
    }
    if (headerRows && headerRows.start === 0 && headerRows.end > 0) sheet.print = { ...(sheet.print ?? DEFAULT_PRINT), repeatRows: headerRows.end };

    // spilled array results: only the formula cell is kept
    for (const { anchor, area } of spillAreas)
      for (let r = area.r1; r <= area.r2; r++)
        for (let c = area.c1; c <= area.c2; c++) {
          const k = cellKey(r, c);
          if (k === anchor) continue;
          const cell = sheet.cells.get(k);
          if (cell && cell.f === undefined) {
            delete cell.v;
            if ((!cell.s || cell.s === defaultId) && !cell.link && !cell.note) sheet.cells.delete(k);
          }
        }
    sheet.markExtentStale();

    // validations
    for (const [key, pending] of validations) {
      const ranges = cellRanges.get(key);
      if (ranges.length) sheet.validations.push({ id: seq(), ranges, ...pending.rule });
    }

    // pictures placed on the sheet rather than in a cell
    for (const sh of shapes)
      for (const fr of kids(sh)) {
        if (fr.localName !== 'frame') continue;
        let x = lengthPx(at(fr, NS.svg, 'x')) ?? 0;
        let y = lengthPx(at(fr, NS.svg, 'y')) ?? 0;
        let c = 0;
        let r = 0;
        while (c < MAX_COLS - 1 && x >= sheet.colWidth(c) && sheet.colWidth(c) > 0) x -= sheet.colWidth(c++);
        while (r < MAX_ROWS - 1 && y >= sheet.rowHeight(r) && sheet.rowHeight(r) > 0) y -= sheet.rowHeight(r++);
        readFrame(fr, sheet, r, c, x, y);
      }

    // conditional formatting
    for (const block of cfBlocks)
      for (const f of kids(block)) {
        if (f.localName !== 'conditional-format') continue;
        const ranges = splitRangeList(at(f, NS.calcext, 'target-range-address') ?? '')
          .map((a) => parseRangeAddress(a)?.range)
          .filter((r): r is Range => !!r);
        if (!ranges.length) continue;
        for (const rule of kids(f)) {
          const base = { id: seq(), ranges: ranges.map((r) => ({ ...r })) };
          switch (rule.localName) {
            case 'condition': {
              const parsed = cfConditionIn(at(rule, NS.calcext, 'value') ?? '');
              if (!parsed) break;
              const styleName = at(rule, NS.calcext, 'apply-style-name') ?? '';
              const style = condStyleIn(reader.byDisplay.get(styleName) ?? reader.styles.get(styleName), reader);
              sheet.cf.push({ ...base, type: 'cell', ...parsed, style } as CondFormat);
              break;
            }
            case 'date-is': {
              const period = DATE_PERIOD_IN[at(rule, NS.calcext, 'date') ?? ''];
              if (!period) break;
              const styleName = at(rule, NS.calcext, 'style') ?? '';
              sheet.cf.push({ ...base, type: 'date', datePeriod: period, style: condStyleIn(reader.byDisplay.get(styleName) ?? reader.styles.get(styleName), reader) });
              break;
            }
            case 'color-scale': {
              const colors = kids(rule)
                .map((e) => colorIn(at(e, NS.calcext, 'color')))
                .filter((c): c is string => !!c);
              if (colors.length >= 2) sheet.cf.push({ ...base, type: 'colorScale', colors: colors.slice(0, 3) });
              break;
            }
            case 'data-bar':
              sheet.cf.push({ ...base, type: 'dataBar', colors: [colorIn(at(rule, NS.calcext, 'positive-color')) ?? '#638ec6'], ...(at(rule, NS.calcext, 'show-value') === 'false' ? { showValue: false } : {}) });
              break;
            case 'icon-set': {
              const name = at(rule, NS.calcext, 'icon-set-type') ?? '3Arrows';
              const known: NonNullable<CondFormat['iconSet']>[] = ['3Arrows', '3TrafficLights', '3Symbols', '3Stars', '4Arrows', '5Arrows', '3Flags', '5Ratings'];
              const set = known.find((k) => name.startsWith(k) || (k === '5Ratings' && name === '5Rating')) ?? (name.startsWith('3') ? '3TrafficLights' : name.startsWith('4') ? '4Arrows' : '5Arrows');
              sheet.cf.push({ ...base, type: 'iconSet', iconSet: set, ...(at(rule, NS.calcext, 'show-value') === 'false' ? { showValue: false } : {}) });
              break;
            }
            default:
              break;
          }
        }
      }
    // files without LibreOffice's extension carry conditions on the cell styles
    if (!cfBlocks.length)
      for (const styleName of legacyCf.keys()) {
        const st = reader.styles.get(styleName);
        const ranges = legacyCf.get(styleName);
        for (const m of st?.maps ?? []) {
          const parsed = styleMapConditionIn(m.cond);
          if (!parsed) continue;
          sheet.cf.push({ id: seq(), ranges: ranges.map((r) => ({ ...r })), type: 'cell', ...parsed, style: condStyleIn(reader.styles.get(m.style) ?? reader.byDisplay.get(m.style), reader) } as CondFormat);
        }
      }

    // sheet-scoped names
    for (const block of localNames) readNames(block, wb, sheet.id);
  }

  // workbook names
  for (const block of kids(spreadsheet)) if (block.localName === 'named-expressions' && block.namespaceURI === NS.table) readNames(block, wb, undefined);

  // filters
  for (const db of kids(child(spreadsheet, NS.table, 'database-ranges') ?? spreadsheet).filter((k) => k.localName === 'database-range')) {
    if (at(db, NS.table, 'display-filter-buttons') !== 'true') continue;
    const target = parseRangeAddress(at(db, NS.table, 'target-range-address') ?? '');
    if (!target) continue;
    const sheet = target.sheet !== undefined ? sheetByName.get(target.sheet) ?? wb.sheetByName(target.sheet) : wb.sheets[0];
    if (!sheet || sheet.filter) continue;
    sheet.filter = { range: target.range, columns: filterColumnsIn(db, target.range.c1) };
  }

  // view settings
  const settings = parse('settings.xml');
  if (settings) readSettings(settings, wb, sheetByName);

  // properties
  // (flat files keep office:meta next to the body, where notes also carry dc:creator)
  const metaRoot = (parse('meta.xml') ?? content).documentElement;
  const metaEl = metaRoot ? child(metaRoot, NS.office, 'meta') : undefined;
  if (metaEl) {
    const text = (ns: string, name: string) => kids(metaEl).find((k) => k.namespaceURI === ns && k.localName === name)?.textContent?.trim() || undefined;
    const created = text(NS.meta, 'creation-date');
    wb.props = {
      title: text(NS.dc, 'title'),
      author: text(NS.meta, 'initial-creator') ?? text(NS.dc, 'creator'),
      created: created && !Number.isNaN(Date.parse(created)) ? Date.parse(created) : undefined,
      company: kids(metaEl).find((u) => u.localName === 'user-defined' && at(u, NS.meta, 'name') === 'Company')?.textContent || undefined,
    };
  }
  if (!wb.sheets.length) wb.addSheet();
  return { wb, warning: warnings.join(' ') || undefined };
}

/** A picture stored inline as base64 (flat OpenDocument files). */
function embeddedPicture(b64: string): string | undefined {
  const mime = b64.startsWith('iVBOR') ? 'image/png' : b64.startsWith('/9j/') ? 'image/jpeg' : b64.startsWith('R0lGOD') ? 'image/gif' : b64.startsWith('PHN2Zy') || b64.startsWith('PD94bW') ? 'image/svg+xml' : undefined;
  return mime ? `data:${mime};base64,${b64}` : undefined;
}

function readNames(block: Element, wb: Workbook, sheetId: number | undefined): void {
  for (const n of kids(block)) {
    const name = at(n, NS.table, 'name');
    if (!name || name.startsWith('__')) continue;
    let ref: string | undefined;
    if (n.localName === 'named-range') {
      const addr = at(n, NS.table, 'cell-range-address');
      if (addr) ref = odsRefToExcel(addr.trim());
    } else if (n.localName === 'named-expression') {
      const expr = at(n, NS.table, 'expression');
      if (expr) ref = fromOdsFormula(expr);
    }
    if (!ref) continue;
    if (wb.names.some((x) => x.name.toLowerCase() === name.toLowerCase() && x.sheet === sheetId)) continue;
    wb.names.push({ name, ref, sheet: sheetId });
  }
}

function filterColumnsIn(db: Element, c1: number): AutoFilter['columns'] {
  const out: AutoFilter['columns'] = {};
  const filter = child(db, NS.table, 'filter');
  if (!filter) return out;
  const conds = descendants(filter, NS.table, 'filter-condition');
  const byField = new Map<number, Element[]>();
  for (const c of conds) {
    const f = Number(at(c, NS.table, 'field-number') ?? -1);
    if (f < 0) continue;
    byField.set(f, [...(byField.get(f) ?? []), c]);
  }
  const ops: Record<string, ColumnFilter['condition'] extends infer T ? (T extends { op: infer O } ? O : never) : never> = {
    '=': 'equal',
    '!=': 'notEqual',
    '>': 'greater',
    '<': 'less',
    '>=': 'greaterEqual',
    '<=': 'lessEqual',
    contains: 'contains',
    '!contains': 'notContains',
    begins: 'begins',
    ends: 'ends',
  };
  for (const [field, list] of byField) {
    const col = c1 + field;
    const first = list[0];
    const items = kids(first).filter((k) => k.localName === 'filter-set-item');
    const op = (at(first, NS.table, 'operator') ?? '=').toLowerCase();
    const value = at(first, NS.table, 'value') ?? '';
    if (items.length) {
      const values = items.map((i) => at(i, NS.table, 'value') ?? '');
      out[col] = { values: values.filter((v) => v !== ''), blanks: values.includes('') || undefined };
    } else if (/^(top|bottom) (values|percent)$/.test(op)) out[col] = { top: { n: Number(value) || 10, bottom: op.startsWith('bottom'), percent: op.endsWith('percent') } };
    else if (list.length === 2 && at(list[0], NS.table, 'operator') === '>=' && at(list[1], NS.table, 'operator') === '<=') out[col] = { condition: { op: 'between', value, value2: at(list[1], NS.table, 'value') ?? '' } };
    else if (op === '=' && list.length === 1) out[col] = { values: [value] };
    else if (ops[op]) out[col] = { condition: { op: ops[op], value } };
  }
  return out;
}

function readSettings(settings: Document, wb: Workbook, sheetByName: Map<string, Sheet>): void {
  const CONFIG = 'urn:oasis:names:tc:opendocument:xmlns:config:1.0';
  const holder = settings.documentElement ? child(settings.documentElement, NS.office, 'settings') : undefined;
  const sets = holder ? kids(holder).filter((k) => k.localName === 'config-item-set') : [];
  const view = sets.find((s) => at(s, CONFIG, 'name') === 'ooo:view-settings');
  if (!view) return;
  const views = kids(view).find((k) => k.localName === 'config-item-map-indexed' && at(k, CONFIG, 'name') === 'Views');
  const entry = views ? kids(views).find((k) => k.localName === 'config-item-map-entry') : undefined;
  if (!entry) return;
  const items = (el: Element) => {
    const out: Record<string, string> = {};
    for (const k of kids(el)) if (k.localName === 'config-item') out[at(k, CONFIG, 'name') ?? ''] = k.textContent ?? '';
    return out;
  };
  const global = items(entry);
  const tables = kids(entry).find((k) => k.localName === 'config-item-map-named' && at(k, CONFIG, 'name') === 'Tables');
  for (const t of tables ? kids(tables) : []) {
    const sheet = sheetByName.get(at(t, CONFIG, 'name') ?? '');
    if (!sheet) continue;
    const it = items(t);
    const num = (k: string) => (it[k] !== undefined ? Number(it[k]) : undefined);
    if (num('HorizontalSplitMode') === 2) sheet.view.freezeCols = Math.max(0, num('HorizontalSplitPosition') ?? 0);
    if (num('VerticalSplitMode') === 2) sheet.view.freezeRows = Math.max(0, num('VerticalSplitPosition') ?? 0);
    const zoom = num('ZoomValue');
    if (zoom) sheet.view.zoom = Math.max(0.25, Math.min(4, zoom / 100));
    if (it.ShowGrid === 'false') sheet.view.showGrid = false;
    const cx = num('CursorPositionX');
    const cy = num('CursorPositionY');
    if (cx !== undefined && cy !== undefined && (cx || cy)) sheet.view.selection = { ranges: [{ r1: cy, c1: cx, r2: cy, c2: cx }], active: { r: cy, c: cx } };
  }
  if (global.ShowGrid === 'false') for (const s of wb.sheets) if (!tables) s.view.showGrid = false;
  if (global.HasColumnRowHeaders === 'false') for (const s of wb.sheets) s.view.showHeaders = false;
  const active = global.ActiveTable ? sheetByName.get(global.ActiveTable) : undefined;
  if (active) wb.activeSheet = Math.max(0, wb.sheets.indexOf(active));
}
