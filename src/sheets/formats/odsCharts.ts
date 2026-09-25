/**
 * Charts in OpenDocument spreadsheets are embedded chart documents ("Object 1/content.xml") that point
 * back at cell ranges. These helpers write a chart document for a Sheets chart and read one back.
 */
import { escapeXml } from '@/lib/utils';
import { chartLayout, PALETTES } from '../charts/chartConfig';
import { formatValue } from '../format/numfmt';
import { isErr } from '../engine/values';
import { colName, MAX_COLS, MAX_ROWS } from '../model/address';
import type { ChartSpec, ChartType } from '../model/types';
import type { Sheet } from '../model/workbook';
import type { SheetDoc } from '../doc';
import { odsRangeAddress, odsRefToExcel, splitRangeList } from './odsFormula';

const NS_CHART = 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0';
const NS_TABLE = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
const NS_STYLE = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';
const NS_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
const NS_DRAW = 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0';
const NS_SVG = 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0';

const CHART_XMLNS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  `xmlns:style="${NS_STYLE}"`,
  `xmlns:text="${NS_TEXT}"`,
  `xmlns:table="${NS_TABLE}"`,
  `xmlns:draw="${NS_DRAW}"`,
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  `xmlns:svg="${NS_SVG}"`,
  `xmlns:chart="${NS_CHART}"`,
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
].join(' ');

const esc = escapeXml;
const inch = (px: number) => `${Number((px / 96).toFixed(4))}in`;

const CLASS_OUT: Record<ChartType, string> = {
  column: 'chart:bar',
  bar: 'chart:bar',
  line: 'chart:line',
  area: 'chart:area',
  pie: 'chart:circle',
  doughnut: 'chart:ring',
  polarArea: 'chart:circle',
  scatter: 'chart:scatter',
  bubble: 'chart:scatter',
  radar: 'chart:radar',
  combo: 'chart:bar',
};

export interface OdsChartObject {
  content: string;
  styles: string;
  /** The ranges the chart reads, for draw:notify-on-update-of-ranges. */
  ranges: string;
}

/** The chart document for a Sheets chart, or null when its source can't be resolved. */
export function odsChartObject(doc: SheetDoc, spec: ChartSpec, home: Sheet): OdsChartObject | null {
  const layout = chartLayout(doc, spec, home);
  if (!layout) return null;
  const { sheet, byRows, headerRow, headerCol, grid } = layout;
  const R = grid.length;
  const C = grid[0]?.length ?? 0;
  const start = headerRow ? 1 : 0;
  const first = headerCol ? 1 : 0;
  if (C - first < 1 || R - start < 1) return null;
  // sheet position of a cell of the grid (rows = points, columns = series)
  const pos = (i: number, j: number) => (byRows ? { r: layout.r1 + j, c: layout.c1 + i } : { r: layout.r1 + i, c: layout.c1 + j });
  const range = (i1: number, j1: number, i2: number, j2: number) => {
    const a = pos(i1, j1);
    const b = pos(i2, j2);
    return odsRangeAddress(sheet.name, Math.min(a.r, b.r), Math.min(a.c, b.c), Math.min(MAX_ROWS - 1, Math.max(a.r, b.r)), Math.min(MAX_COLS - 1, Math.max(a.c, b.c)));
  };
  const text = (i: number, j: number) => {
    const v = grid[i]?.[j];
    if (v === null || v === undefined || isErr(v)) return '';
    const p = pos(i, j);
    return formatValue(v, doc.styleOf(sheet, p.r, p.c).numFmt).text;
  };
  const type = spec.type;
  const pal = PALETTES[spec.palette ?? 'affice'] ?? PALETTES.affice;
  const color = (i: number) => pal[i % pal.length];
  const pie = type === 'pie' || type === 'doughnut' || type === 'polarArea';
  const xy = type === 'scatter' || type === 'bubble';
  const cats = headerCol ? range(start, 0, R - 1, 0) : undefined;
  const points = R - start;
  const seriesCount = C - first;

  // styles
  const styles: string[] = [];
  const style = (name: string, body: string) => styles.push(`<style:style style:name="${name}" style:family="chart">${body}</style:style>`);
  style('ch1', '<style:graphic-properties draw:stroke="none" draw:fill="solid" draw:fill-color="#ffffff"/>');
  style('ch2', '<style:chart-properties chart:auto-position="true"/><style:text-properties fo:font-size="13pt" fo:font-weight="bold" style:font-size-asian="13pt" style:font-size-complex="13pt"/>');
  style('ch3', '<style:chart-properties chart:auto-position="true"/><style:text-properties fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>');
  const plotProps = [
    'chart:auto-position="true"',
    'chart:auto-size="true"',
    'chart:treat-empty-cells="leave-gap"',
    spec.stacked === 'percent' ? 'chart:percentage="true"' : spec.stacked ? 'chart:stacked="true"' : '',
    type === 'bar' ? 'chart:vertical="true"' : '',
    spec.smooth && (type === 'line' || xy || type === 'combo') ? 'chart:interpolation="cubic-spline"' : '',
    type === 'doughnut' ? 'chart:hole-size="55%"' : '',
  ].filter(Boolean);
  style('ch4', `<style:chart-properties ${plotProps.join(' ')}/>`);
  style('ch5', '<style:chart-properties chart:display-label="true" chart:link-data-style-to-source="true"/><style:graphic-properties svg:stroke-color="#b3b3b3"/><style:text-properties fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>');
  style('ch6', '<style:graphic-properties svg:stroke-color="#d9d9d9"/>');
  style('ch7', '<style:graphic-properties draw:stroke="none" draw:fill="none"/>');
  style('ch8', '<style:chart-properties chart:auto-position="true"/><style:text-properties fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>');
  const labels = spec.dataLabels ? ' chart:data-label-number="value"' : '';
  const seriesClass = (k: number): string => (type === 'combo' ? (k === 0 ? 'chart:bar' : 'chart:line') : CLASS_OUT[type] ?? 'chart:bar');
  for (let k = 0; k < seriesCount; k++) {
    const cls = seriesClass(k);
    const line = cls === 'chart:line' || cls === 'chart:scatter' || cls === 'chart:radar';
    // like the on-screen chart: points on lines unless there are very many
    const markers = cls === 'chart:scatter' || cls === 'chart:radar' || (cls === 'chart:line' && points <= 60);
    const symbol = markers ? ' chart:symbol-type="named-symbol" chart:symbol-name="circle" chart:symbol-width="0.2cm" chart:symbol-height="0.2cm"' : ' chart:symbol-type="none"';
    if (line) {
      const stroke = cls === 'chart:scatter' && !spec.smooth ? 'draw:stroke="none"' : `svg:stroke-width="0.07cm" svg:stroke-color="${color(k)}"`;
      style(`chs${k}`, `<style:chart-properties${symbol}${labels} chart:link-data-style-to-source="true"/><style:graphic-properties ${stroke} draw:fill-color="${color(k)}"/><style:text-properties fo:font-size="10pt"/>`);
    } else style(`chs${k}`, `<style:chart-properties${labels} chart:link-data-style-to-source="true"/><style:graphic-properties draw:stroke="none" draw:fill-color="${color(k)}"/><style:text-properties fo:font-size="10pt"/>`);
  }
  if (pie) for (let i = 0; i < points; i++) style(`chp${i}`, `<style:graphic-properties draw:stroke="solid" svg:stroke-color="#ffffff" draw:fill-color="${color(i)}"/>`);

  // plot area
  const horizontal = type === 'bar';
  const axisTitle = (t?: string) => (t ? `<chart:title chart:style-name="ch8"><text:p>${esc(t)}</text:p></chart:title>` : '');
  let axes = '';
  if (!pie) {
    axes += `<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ch5">${axisTitle(horizontal ? spec.yTitle : spec.xTitle)}${cats ? `<chart:categories table:cell-range-address="${esc(cats)}"/>` : ''}</chart:axis>`;
    axes += `<chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch5">${axisTitle(horizontal ? spec.xTitle : spec.yTitle)}${spec.gridlines !== false ? '<chart:grid chart:style-name="ch6" chart:class="major"/>' : ''}</chart:axis>`;
    if (type === 'combo' && seriesCount === 2) axes += '<chart:axis chart:dimension="y" chart:name="secondary-y" chart:style-name="ch5"/>';
  } else if (cats) axes += `<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ch5"><chart:categories table:cell-range-address="${esc(cats)}"/></chart:axis><chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch5"/>`;
  let seriesXml = '';
  const used: string[] = [];
  if (cats) used.push(cats);
  for (let k = 0; k < seriesCount; k++) {
    const j = first + k;
    const values = range(start, j, R - 1, j);
    const label = headerRow ? range(0, j, 0, j) : undefined;
    used.push(...(label ? [label] : []), values);
    const attrs = [`chart:style-name="chs${k}"`, `chart:values-cell-range-address="${esc(values)}"`];
    if (label) attrs.push(`chart:label-cell-address="${esc(label)}"`);
    attrs.push(`chart:class="${seriesClass(k)}"`);
    if (type === 'combo' && seriesCount === 2 && k === 1) attrs.push('chart:attached-axis="secondary-y"');
    const domain = xy && cats ? `<chart:domain table:cell-range-address="${esc(cats)}"/>` : '';
    const pts = pie ? Array.from({ length: points }, (_, i) => `<chart:data-point chart:style-name="chp${i}"/>`).join('') : `<chart:data-point chart:repeated="${points}"/>`;
    seriesXml += `<chart:series ${attrs.join(' ')}>${domain}${pts}</chart:series>`;
    if (pie) break; // a pie shows one series
  }
  const src = odsRangeAddress(sheet.name, layout.r1, layout.c1, layout.r2, layout.c2);
  const hasLabels = (() => {
    const firstRow = byRows ? headerCol : headerRow;
    const firstCol = byRows ? headerRow : headerCol;
    return firstRow && firstCol ? 'both' : firstRow ? 'row' : firstCol ? 'column' : 'none';
  })();
  const plot = `<chart:plot-area chart:style-name="ch4" table:cell-range-address="${esc(src)}" chart:data-source-has-labels="${hasLabels}">${axes}${seriesXml}${pie ? '' : '<chart:wall chart:style-name="ch7"/>'}</chart:plot-area>`;
  const legendPos = spec.legend === 'top' ? 'top' : spec.legend === 'left' ? 'start' : spec.legend === 'right' ? 'end' : 'bottom';
  const legend = spec.legend === 'none' ? '' : `<chart:legend chart:legend-position="${legendPos}" chart:style-name="ch3"/>`;
  const title = spec.title ? `<chart:title chart:style-name="ch2"><text:p>${esc(spec.title)}</text:p></chart:title>` : '';

  // the chart's own copy of the data, for apps that don't read the cells
  const shown = pie ? 1 : seriesCount;
  let local = `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns><table:table-column table:number-columns-repeated="${shown}"/></table:table-columns>`;
  local += '<table:table-header-rows><table:table-row><table:table-cell><text:p/></table:table-cell>';
  for (let k = 0; k < shown; k++) {
    const name = headerRow ? text(0, first + k) : '';
    local += `<table:table-cell office:value-type="string"><text:p>${esc(name || `Series ${k + 1}`)}</text:p></table:table-cell>`;
  }
  local += '</table:table-row></table:table-header-rows><table:table-rows>';
  for (let i = start; i < R; i++) {
    const cat = headerCol ? text(i, 0) : String(i - start + 1);
    local += `<table:table-row><table:table-cell office:value-type="string"><text:p>${esc(cat)}</text:p></table:table-cell>`;
    for (let k = 0; k < shown; k++) {
      const v = grid[i][first + k];
      local += typeof v === 'number' && Number.isFinite(v) ? `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>` : '<table:table-cell><text:p/></table:table-cell>';
    }
    local += '</table:table-row>';
  }
  local += '</table:table-rows></table:table>';

  const size = `svg:width="${inch(spec.w)}" svg:height="${inch(spec.h)}"`;
  const content = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${CHART_XMLNS} office:version="1.3"><office:automatic-styles>${styles.join('')}</office:automatic-styles><office:body><office:chart><chart:chart ${size} xlink:href=".." xlink:type="simple" chart:class="${CLASS_OUT[type] ?? 'chart:bar'}" chart:style-name="ch1">${title}${legend}${plot}${local}</chart:chart></office:chart></office:body></office:document-content>`;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${CHART_XMLNS} office:version="1.3"><office:styles/></office:document-styles>`;
  return { content, styles: stylesXml, ranges: used.join(' ') };
}

/* ================================================================ reading */

const TYPE_IN: Record<string, ChartType> = {
  'chart:bar': 'column',
  'chart:line': 'line',
  'chart:area': 'area',
  'chart:circle': 'pie',
  'chart:ring': 'doughnut',
  'chart:scatter': 'scatter',
  'chart:bubble': 'bubble',
  'chart:radar': 'radar',
  'chart:filled-radar': 'radar',
  'chart:stock': 'line',
};

function kids(el: Element): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
  return out;
}

const at = (el: Element | undefined, ns: string, name: string) => (el && el.hasAttributeNS(ns, name) ? el.getAttributeNS(ns, name) ?? undefined : undefined);

function textIn(el: Element | undefined): string | undefined {
  if (!el) return undefined;
  const paras = kids(el).filter((k) => k.namespaceURI === NS_TEXT && k.localName === 'p');
  const t = (paras.length ? paras.map((p) => p.textContent ?? '').join(' ') : el.textContent ?? '').trim();
  return t || undefined;
}

interface Box {
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function boxOf(addr: string | undefined): Box[] {
  if (!addr) return [];
  const out: Box[] = [];
  for (const part of splitRangeList(addr)) {
    const ex = odsRefToExcel(part);
    const m = /^(?:('(?:[^']|'')+'|[^!]+)!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(ex);
    if (!m) continue;
    const sheet = m[1] ? (m[1].startsWith("'") ? m[1].slice(1, -1).replace(/''/g, "'") : m[1]) : '';
    const col = (s: string) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    const r1 = Number(m[3]) - 1;
    const c1 = col(m[2]);
    const r2 = m[5] ? Number(m[5]) - 1 : r1;
    const c2 = m[4] ? col(m[4]) : c1;
    out.push({ sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) });
  }
  return out;
}

/** Reads a chart document into the chart model (without its position, which the frame gives). */
export function readOdsChart(root: Document | Element, homeSheet: string): Omit<ChartSpec, 'id' | 'anchor' | 'w' | 'h'> | null {
  const chart = Array.from(root.getElementsByTagNameNS(NS_CHART, 'chart'))[0];
  if (!chart) return null;
  const styles = new Map<string, { chart: Element | undefined; graphic: Element | undefined }>();
  for (const s of Array.from(root.getElementsByTagNameNS(NS_STYLE, 'style'))) {
    const name = at(s, NS_STYLE, 'name');
    if (!name) continue;
    const props = kids(s);
    styles.set(name, { chart: props.find((p) => p.localName === 'chart-properties'), graphic: props.find((p) => p.localName === 'graphic-properties') });
  }
  const styleOf = (el: Element | undefined) => styles.get(at(el, NS_CHART, 'style-name') ?? '');
  const plot = kids(chart).find((k) => k.localName === 'plot-area' && k.namespaceURI === NS_CHART);
  if (!plot) return null;
  const plotProps = styleOf(plot)?.chart;
  const cls = at(chart, NS_CHART, 'class') ?? 'chart:bar';
  let type: ChartType = TYPE_IN[cls] ?? 'column';
  const series = kids(plot).filter((k) => k.localName === 'series' && k.namespaceURI === NS_CHART);
  if (!series.length) return null;
  const classes = series.map((s) => at(s, NS_CHART, 'class') ?? cls);
  if (type === 'column' && at(plotProps, NS_CHART, 'vertical') === 'true') type = 'bar';
  if ((type === 'column' || type === 'bar') && classes.slice(1).some((c) => c === 'chart:line')) type = 'combo';
  const axes = kids(plot).filter((k) => k.localName === 'axis' && k.namespaceURI === NS_CHART);
  const axis = (dim: string) => axes.find((a) => at(a, NS_CHART, 'dimension') === dim && (at(a, NS_CHART, 'name') ?? '').startsWith('primary')) ?? axes.find((a) => at(a, NS_CHART, 'dimension') === dim);
  const xAxis = axis('x');
  const yAxis = axis('y');
  const title = (el: Element | undefined) => textIn(el ? kids(el).find((k) => k.localName === 'title' && k.namespaceURI === NS_CHART) : undefined);
  const catAddr = xAxis ? at(kids(xAxis).find((k) => k.localName === 'categories'), NS_TABLE, 'cell-range-address') : undefined;

  // the source: every range the chart reads, on the first sheet it names
  const boxes: Box[] = [...boxOf(at(plot, NS_TABLE, 'cell-range-address'))];
  if (!boxes.length) {
    boxes.push(...boxOf(catAddr));
    for (const s of series) boxes.push(...boxOf(at(s, NS_CHART, 'values-cell-range-address')), ...boxOf(at(s, NS_CHART, 'label-cell-address')));
    for (const d of Array.from(plot.getElementsByTagNameNS(NS_CHART, 'domain'))) boxes.push(...boxOf(at(d, NS_TABLE, 'cell-range-address')));
  }
  if (!boxes.length) return null;
  const sheet = boxes[0].sheet || homeSheet;
  const same = boxes.filter((b) => (b.sheet || homeSheet) === sheet);
  const box = { r1: Math.min(...same.map((b) => b.r1)), c1: Math.min(...same.map((b) => b.c1)), r2: Math.max(...same.map((b) => b.r2)), c2: Math.max(...same.map((b) => b.c2)) };
  const values = series.map((s) => boxOf(at(s, NS_CHART, 'values-cell-range-address'))[0]).filter((b): b is Box => !!b);
  const seriesInRows = values.length > 0 && values.every((b) => b.r1 === b.r2 && b.c2 > b.c1);

  // colours → the closest palette
  const colors = series
    .flatMap((s) => {
      const pts = kids(s).filter((k) => k.localName === 'data-point');
      const own = at(styleOf(s)?.graphic, NS_DRAW, 'fill-color');
      return type === 'pie' || type === 'doughnut' ? pts.map((p) => at(styleOf(p)?.graphic, NS_DRAW, 'fill-color')).filter(Boolean) : [own ?? at(styleOf(s)?.graphic, NS_SVG, 'stroke-color')];
    })
    .filter((c): c is string => !!c)
    .map((c) => c.toLowerCase());
  const palette = Object.entries(PALETTES).find(([, pal]) => colors.length > 0 && colors.every((c, i) => pal[i % pal.length].toLowerCase() === c))?.[0] ?? 'office';

  const legend = kids(chart).find((k) => k.localName === 'legend' && k.namespaceURI === NS_CHART);
  const legendPos = at(legend, NS_CHART, 'legend-position');
  const interpolation = at(plotProps, NS_CHART, 'interpolation');
  const labelsShown = series.some((s) => {
    const v = at(styleOf(s)?.chart, NS_CHART, 'data-label-number');
    return v !== undefined && v !== 'none';
  });
  const horizontal = type === 'bar';
  const quote = /^[A-Za-z_][\w.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  return {
    type,
    source: `${quote}!$${colName(box.c1)}$${box.r1 + 1}:$${colName(box.c2)}$${box.r2 + 1}`,
    seriesInRows,
    stacked: at(plotProps, NS_CHART, 'percentage') === 'true' ? 'percent' : at(plotProps, NS_CHART, 'stacked') === 'true' ? true : undefined,
    title: textIn(kids(chart).find((k) => k.localName === 'title' && k.namespaceURI === NS_CHART)),
    xTitle: horizontal ? title(yAxis) : title(xAxis),
    yTitle: horizontal ? title(xAxis) : title(yAxis),
    legend: !legend ? 'none' : legendPos === 'top' ? 'top' : legendPos === 'start' ? 'left' : legendPos === 'end' ? 'right' : 'bottom',
    palette,
    smooth: (!!interpolation && interpolation !== 'none') || undefined,
    dataLabels: labelsShown || undefined,
    gridlines: yAxis && !kids(yAxis).some((k) => k.localName === 'grid') && type !== 'pie' && type !== 'doughnut' ? false : undefined,
  };
}
