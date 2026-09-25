/**
 * Slide charts in OpenDocument presentations: an embedded chart document ("Object N") that keeps its data
 * in a local table, like LibreOffice Impress does, so the chart can be edited without a spreadsheet.
 */
import { escapeXml } from '@/lib/utils';
import { resolveHex, type SlideChart, type SlideChartType, type Theme } from '../model';

const NS_CHART = 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0';
const NS_TABLE = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
const NS_STYLE = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';
const NS_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
const NS_DRAW = 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0';
const NS_SVG = 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0';
const NS_OFFICE = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0';
const NS_FO = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0';

const XMLNS = [
  `xmlns:office="${NS_OFFICE}"`,
  `xmlns:style="${NS_STYLE}"`,
  `xmlns:text="${NS_TEXT}"`,
  `xmlns:table="${NS_TABLE}"`,
  `xmlns:draw="${NS_DRAW}"`,
  `xmlns:fo="${NS_FO}"`,
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  `xmlns:svg="${NS_SVG}"`,
  `xmlns:chart="${NS_CHART}"`,
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
].join(' ');

const esc = escapeXml;
const cm = (px: number) => `${Math.round((px * 2.54 * 1000) / 96) / 1000}cm`;

const CLASS: Record<SlideChartType, string> = {
  column: 'chart:bar',
  bar: 'chart:bar',
  line: 'chart:line',
  area: 'chart:area',
  pie: 'chart:circle',
  doughnut: 'chart:ring',
  scatter: 'chart:scatter',
  radar: 'chart:radar',
};

const colLetter = (i: number) => {
  let s = '';
  let n = i + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

/** Mixes two #rrggbb colours (t = share of b). */
function mix(a: string, b: string, t: number): string {
  const p = (h: string, i: number) => parseInt(h.slice(1 + i * 2, 3 + i * 2), 16);
  const c = (i: number) =>
    Math.round(p(a, i) * (1 - t) + p(b, i) * t)
      .toString(16)
      .padStart(2, '0');
  return `#${c(0)}${c(1)}${c(2)}`;
}

/** The colour of series i (or pie slice i), as on screen: theme accents, then lighter repeats. */
export function chartSeriesHex(chart: SlideChart, i: number, theme: Theme): string {
  const own = chart.series[i]?.color;
  if (own) return resolveHex(own, theme).hex;
  const base = resolveHex(`@accent${(i % 6) + 1}`, theme).hex;
  return i < 6 ? base : mix(base, resolveHex('@bg1', theme, '#ffffff').hex, 0.4);
}

/** The chart document (content.xml and styles.xml of "Object N") for a slide chart. */
export function slideChartObject(chart: SlideChart, theme: Theme, w: number, h: number): { content: string; styles: string } {
  const type = chart.type;
  const pie = type === 'pie' || type === 'doughnut';
  const xy = type === 'scatter';
  const text = resolveHex('@tx1', theme).hex;
  const bg = resolveHex('@bg1', theme, '#ffffff').hex;
  const grid = mix(bg, text, 0.12);
  const font = theme.fonts.minor;
  const size = chart.fontSize ?? 14;
  const titleSize = chart.titleSize ?? Math.round(size * 1.29 * 10) / 10;
  const series = pie ? chart.series.slice(0, 1) : chart.series;
  const n = chart.categories.length;
  const textProps = (pt: number, bold = false, color = text) =>
    `<style:text-properties fo:color="${color}" fo:font-family="${esc(quoteFont(font))}" fo:font-size="${pt}pt" style:font-size-asian="${pt}pt" style:font-size-complex="${pt}pt"${bold ? ' fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"' : ''}/>`;

  const styles: string[] = [];
  const style = (name: string, body: string) => styles.push(`<style:style style:name="${name}" style:family="chart">${body}</style:style>`);
  style('ch1', '<style:graphic-properties draw:stroke="none" draw:fill="none"/>');
  style('ch2', `<style:chart-properties chart:auto-position="true"/>${textProps(titleSize, true)}`);
  style('ch3', `<style:chart-properties chart:auto-position="true"/>${textProps(size, false, mix(bg, text, 0.75))}`);
  const plotProps = [
    'chart:auto-position="true"',
    'chart:auto-size="true"',
    'chart:treat-empty-cells="leave-gap"',
    chart.stacked === 'percent' ? 'chart:percentage="true"' : chart.stacked ? 'chart:stacked="true"' : '',
    type === 'bar' ? 'chart:vertical="true"' : '',
    chart.smooth && (type === 'line' || xy) ? 'chart:interpolation="cubic-spline"' : '',
    type === 'doughnut' ? 'chart:hole-size="55%"' : '',
    type === 'column' || type === 'bar' ? `chart:gap-width="80" chart:overlap="${chart.stacked ? 100 : -10}"` : '',
  ].filter(Boolean);
  style('ch4', `<style:chart-properties ${plotProps.join(' ')}/>`);
  style('ch5', `<style:chart-properties chart:display-label="true" chart:tick-marks-major-inner="false" chart:tick-marks-major-outer="false"/><style:graphic-properties svg:stroke-color="${mix(bg, text, 0.35)}"/>${textProps(size, false, mix(bg, text, 0.65))}`);
  style('ch6', `<style:graphic-properties svg:stroke-color="${grid}"/>`);
  style('ch7', '<style:graphic-properties draw:stroke="none" draw:fill="none"/>');
  const labels = chart.dataLabels ? ' chart:data-label-number="value"' : '';
  series.forEach((_, k) => {
    const color = chartSeriesHex(chart, k, theme);
    if (type === 'line' || type === 'radar') {
      style(`chs${k}`, `<style:chart-properties chart:symbol-type="named-symbol" chart:symbol-name="circle" chart:symbol-width="0.2cm" chart:symbol-height="0.2cm"${labels} chart:link-data-style-to-source="true"/><style:graphic-properties svg:stroke-width="0.079cm" svg:stroke-color="${color}" draw:fill-color="${color}"/>${textProps(size)}`);
    } else if (xy) {
      style(`chs${k}`, `<style:chart-properties chart:symbol-type="named-symbol" chart:symbol-name="circle" chart:symbol-width="0.25cm" chart:symbol-height="0.25cm"${labels} chart:link-data-style-to-source="true"/><style:graphic-properties draw:stroke="none" svg:stroke-color="${color}" draw:fill-color="${color}"/>${textProps(size)}`);
    } else if (type === 'area') {
      style(`chs${k}`, `<style:chart-properties${labels} chart:link-data-style-to-source="true"/><style:graphic-properties draw:stroke="solid" svg:stroke-width="0.053cm" svg:stroke-color="${color}" draw:fill-color="${color}" draw:opacity="65%"/>${textProps(size)}`);
    } else {
      style(`chs${k}`, `<style:chart-properties${labels} chart:link-data-style-to-source="true"/><style:graphic-properties draw:stroke="none" draw:fill-color="${color}"/>${textProps(size)}`);
    }
  });
  if (pie) for (let i = 0; i < n; i++) style(`chp${i}`, `<style:graphic-properties draw:stroke="solid" svg:stroke-width="0.053cm" svg:stroke-color="${bg}" draw:fill-color="${chartSeriesHex({ ...chart, series: [] }, i, theme)}"/>`);

  // plot area: data lives in the local table (A = categories, B… = series)
  const last = n + 1;
  const cats = `local-table.$A$2:.$A$${last}`;
  let axes = '';
  if (!pie) {
    axes += `<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ch5"><chart:categories table:cell-range-address="${cats}"/></chart:axis>`;
    axes += `<chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch5">${chart.gridlines === false ? '' : '<chart:grid chart:style-name="ch6" chart:class="major"/>'}</chart:axis>`;
  } else axes += `<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ch5"><chart:categories table:cell-range-address="${cats}"/></chart:axis><chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch5"/>`;
  let seriesXml = '';
  series.forEach((_, k) => {
    const col = colLetter(k + 1);
    const attrs = [`chart:style-name="chs${k}"`, `chart:values-cell-range-address="local-table.$${col}$2:.$${col}$${last}"`, `chart:label-cell-address="local-table.$${col}$1"`, `chart:class="${CLASS[type]}"`];
    const domain = xy ? `<chart:domain table:cell-range-address="${cats}"/>` : '';
    const pts = pie ? Array.from({ length: n }, (_, i) => `<chart:data-point chart:style-name="chp${i}"/>`).join('') : `<chart:data-point chart:repeated="${Math.max(1, n)}"/>`;
    seriesXml += `<chart:series ${attrs.join(' ')}>${domain}${pts}</chart:series>`;
  });
  const plot = `<chart:plot-area chart:style-name="ch4" table:cell-range-address="local-table.$A$1:.$${colLetter(series.length)}$${last}" chart:data-source-has-labels="both">${axes}${seriesXml}${pie ? '' : '<chart:wall chart:style-name="ch7"/>'}</chart:plot-area>`;
  // like the renderer: a legend when asked, or by default for several series and for pies
  const legendPos = chart.legend ?? (chart.series.length > 1 || pie ? 'bottom' : 'none');
  const legend = legendPos === 'none' ? '' : `<chart:legend chart:legend-position="${legendPos === 'top' ? 'top' : legendPos === 'left' ? 'start' : legendPos === 'right' ? 'end' : 'bottom'}" chart:style-name="ch3"/>`;
  const title = chart.title ? `<chart:title chart:style-name="ch2"><text:p>${esc(chart.title)}</text:p></chart:title>` : '';

  // the local table
  let local = `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns><table:table-column table:number-columns-repeated="${Math.max(1, series.length)}"/></table:table-columns>`;
  local += '<table:table-header-rows><table:table-row><table:table-cell><text:p/></table:table-cell>';
  for (const s of series) local += `<table:table-cell office:value-type="string"><text:p>${esc(s.name)}</text:p></table:table-cell>`;
  local += '</table:table-row></table:table-header-rows><table:table-rows>';
  chart.categories.forEach((c, i) => {
    const num = Number(c);
    const cat = xy && c.trim() !== '' && Number.isFinite(num) ? `<table:table-cell office:value-type="float" office:value="${num}"><text:p>${esc(c)}</text:p></table:table-cell>` : `<table:table-cell office:value-type="string"><text:p>${esc(c)}</text:p></table:table-cell>`;
    local += `<table:table-row>${cat}`;
    for (const s of series) {
      const v = s.values[i];
      local += typeof v === 'number' && Number.isFinite(v) ? `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>` : '<table:table-cell><text:p/></table:table-cell>';
    }
    local += '</table:table-row>';
  });
  local += '</table:table-rows></table:table>';

  const content = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${XMLNS} office:version="1.3"><office:automatic-styles>${styles.join('')}</office:automatic-styles><office:body><office:chart><chart:chart svg:width="${cm(w)}" svg:height="${cm(h)}" xlink:href="." xlink:type="simple" chart:class="${CLASS[type]}" chart:style-name="ch1">${title}${legend}${plot}${local}</chart:chart></office:chart></office:body></office:document-content>`;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${XMLNS} office:version="1.3"><office:styles/></office:document-styles>`;
  return { content, styles: stylesXml };
}

function quoteFont(f: string): string {
  return /[\s,]/.test(f) ? `'${f}'` : f;
}

/* ================================================================ reading */

const TYPE_IN: Record<string, SlideChartType> = {
  'chart:bar': 'column',
  'chart:line': 'line',
  'chart:area': 'area',
  'chart:circle': 'pie',
  'chart:ring': 'doughnut',
  'chart:scatter': 'scatter',
  'chart:bubble': 'scatter',
  'chart:radar': 'radar',
  'chart:filled-radar': 'radar',
  'chart:stock': 'line',
};

function kids(el: Element | undefined): Element[] {
  const out: Element[] = [];
  if (!el) return out;
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
  return out;
}

const kid = (el: Element | undefined, ns: string, name: string) => kids(el).find((k) => k.namespaceURI === ns && k.localName === name);
const at = (el: Element | undefined, ns: string, name: string) => (el && el.hasAttributeNS(ns, name) ? (el.getAttributeNS(ns, name) ?? undefined) : undefined);

function textOf(el: Element | undefined): string {
  if (!el) return '';
  const paras = kids(el).filter((k) => k.namespaceURI === NS_TEXT && (k.localName === 'p' || k.localName === 'h'));
  return (paras.length ? paras.map((p) => p.textContent ?? '').join(' ') : (el.textContent ?? '')).trim();
}

interface Ref {
  c1: number;
  r1: number;
  c2: number;
  r2: number;
}

/** Parses "local-table.$B$2:.$B$7" (or any table name) into zero-based cell bounds. */
function refOf(addr: string | undefined): Ref | null {
  if (!addr) return null;
  const first = addr.trim().split(/\s+/)[0];
  const m = /^(?:'[^']*'|[^.:]*)?\.\$?([A-Za-z]+)\$?(\d+)(?::(?:'[^']*'|[^.:]*)?\.\$?([A-Za-z]+)\$?(\d+))?$/.exec(first);
  if (!m) return null;
  const col = (s: string) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const c1 = col(m[1]);
  const r1 = Number(m[2]) - 1;
  const c2 = m[3] ? col(m[3]) : c1;
  const r2 = m[4] ? Number(m[4]) - 1 : r1;
  return { c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2) };
}

/** Cells of the chart's local table as a grid (strings or numbers). */
function localGrid(table: Element | undefined): (string | number | null)[][] {
  const grid: (string | number | null)[][] = [];
  if (!table) return grid;
  const rows: Element[] = [];
  const walk = (el: Element) => {
    for (const k of kids(el)) {
      if (k.namespaceURI !== NS_TABLE) continue;
      if (k.localName === 'table-row') rows.push(k);
      else if (k.localName === 'table-header-rows' || k.localName === 'table-rows' || k.localName === 'table-row-group') walk(k);
    }
  };
  walk(table);
  for (const row of rows) {
    const out: (string | number | null)[] = [];
    for (const cell of kids(row)) {
      if (cell.namespaceURI !== NS_TABLE || (cell.localName !== 'table-cell' && cell.localName !== 'covered-table-cell')) continue;
      const rep = Math.min(1000, Number(at(cell, NS_TABLE, 'number-columns-repeated') ?? 1) || 1);
      const type = at(cell, NS_OFFICE, 'value-type');
      const v = type === 'float' || type === 'percentage' || type === 'currency' ? Number(at(cell, NS_OFFICE, 'value')) : null;
      const t = textOf(cell);
      for (let i = 0; i < rep; i++) out.push(v !== null && Number.isFinite(v) ? v : t === '' ? null : t);
    }
    const rep = Math.min(1000, Number(at(row, NS_TABLE, 'number-rows-repeated') ?? 1) || 1);
    for (let i = 0; i < rep; i++) grid.push(out.slice());
  }
  return grid;
}

function slice(grid: (string | number | null)[][], ref: Ref | null): (string | number | null)[] {
  if (!ref) return [];
  const out: (string | number | null)[] = [];
  const byRow = ref.r1 === ref.r2 && ref.c2 > ref.c1;
  if (byRow) for (let c = ref.c1; c <= ref.c2; c++) out.push(grid[ref.r1]?.[c] ?? null);
  else for (let r = ref.r1; r <= ref.r2; r++) out.push(grid[r]?.[ref.c1] ?? null);
  return out;
}

/** Reads a chart document (content root of "Object N") into a slide chart. */
export function readSlideChart(root: Document | Element, theme: Theme): SlideChart | null {
  const chartEl = Array.from(root.getElementsByTagNameNS(NS_CHART, 'chart'))[0];
  if (!chartEl) return null;
  const styles = new Map<string, Element>();
  for (const s of Array.from(root.getElementsByTagNameNS(NS_STYLE, 'style'))) {
    const name = at(s, NS_STYLE, 'name');
    if (name) styles.set(name, s);
  }
  const propsOf = (el: Element | undefined, kind: string) => kid(styles.get(at(el, NS_CHART, 'style-name') ?? ''), NS_STYLE, kind);
  const plot = kid(chartEl, NS_CHART, 'plot-area');
  if (!plot) return null;
  const plotProps = propsOf(plot, 'chart-properties');
  const cls = at(chartEl, NS_CHART, 'class') ?? 'chart:bar';
  let type: SlideChartType = TYPE_IN[cls] ?? 'column';
  if (type === 'column' && at(plotProps, NS_CHART, 'vertical') === 'true') type = 'bar';
  const grid = localGrid(Array.from(chartEl.getElementsByTagNameNS(NS_TABLE, 'table'))[0]);
  const seriesEls = kids(plot).filter((k) => k.namespaceURI === NS_CHART && k.localName === 'series');
  const axes = kids(plot).filter((k) => k.namespaceURI === NS_CHART && k.localName === 'axis');
  const xAxis = axes.find((a) => at(a, NS_CHART, 'dimension') === 'x');
  const yAxis = axes.find((a) => at(a, NS_CHART, 'dimension') === 'y');
  const catEl = kid(xAxis, NS_CHART, 'categories') ?? Array.from(plot.getElementsByTagNameNS(NS_CHART, 'domain'))[0];
  let categories = slice(grid, refOf(at(catEl, NS_TABLE, 'cell-range-address'))).map((c) => (c === null ? '' : String(c)));
  const series: SlideChart['series'] = [];
  const hasLabels = at(plot, NS_CHART, 'data-source-has-labels') ?? 'none';
  const whole = refOf(at(plot, NS_TABLE, 'cell-range-address'));
  if (seriesEls.length) {
    seriesEls.forEach((s, k) => {
      const values = slice(grid, refOf(at(s, NS_CHART, 'values-cell-range-address'))).map((v) => (typeof v === 'number' ? v : v === null || v === '' || Number.isNaN(Number(v)) ? null : Number(v)));
      const labelRef = refOf(at(s, NS_CHART, 'label-cell-address'));
      const name = labelRef ? String(grid[labelRef.r1]?.[labelRef.c1] ?? '') : `Series ${k + 1}`;
      const fill = at(propsOf(s, 'graphic-properties'), NS_DRAW, 'fill-color') ?? at(propsOf(s, 'graphic-properties'), NS_SVG, 'stroke-color');
      series.push({ name, values, ...(fill ? { color: fill.toLowerCase() } : {}) });
    });
  } else if (whole) {
    // no series elements: every column after the first holds a series
    const top = hasLabels === 'both' || hasLabels === 'row' ? whole.r1 + 1 : whole.r1;
    const left = hasLabels === 'both' || hasLabels === 'column' ? whole.c1 + 1 : whole.c1;
    if (!categories.length && left > whole.c1) for (let r = top; r <= whole.r2; r++) categories.push(String(grid[r]?.[whole.c1] ?? ''));
    for (let c = left; c <= whole.c2; c++) {
      const values: (number | null)[] = [];
      for (let r = top; r <= whole.r2; r++) {
        const v = grid[r]?.[c];
        values.push(typeof v === 'number' ? v : null);
      }
      series.push({ name: top > whole.r1 ? String(grid[whole.r1]?.[c] ?? '') : `Series ${c - left + 1}`, values });
    }
  }
  if (!series.length) return null;
  const count = Math.max(categories.length, ...series.map((s) => s.values.length));
  if (!categories.length) categories = Array.from({ length: count }, (_, i) => String(i + 1));
  const out: SlideChart = { type, categories, series: series.map((s) => ({ ...s, values: categories.map((_, i) => s.values[i] ?? null) })) };
  // colours that match the default theme colours are left implicit so they follow theme changes
  out.series = out.series.map((s, i) => {
    if (!s.color) return s;
    const def = chartSeriesHex({ ...out, series: out.series.map((x) => ({ ...x, color: undefined })) }, i, theme);
    const { color, ...rest } = s;
    return color.toLowerCase() === def.toLowerCase() || type === 'pie' || type === 'doughnut' ? rest : { ...rest, color };
  });
  if (at(plotProps, NS_CHART, 'percentage') === 'true') out.stacked = 'percent';
  else if (at(plotProps, NS_CHART, 'stacked') === 'true') out.stacked = true;
  const title = textOf(kid(chartEl, NS_CHART, 'title'));
  if (title) out.title = title;
  const legend = kid(chartEl, NS_CHART, 'legend');
  const pos = at(legend, NS_CHART, 'legend-position');
  out.legend = !legend ? 'none' : pos === 'top' ? 'top' : pos === 'start' ? 'left' : pos === 'end' ? 'right' : 'bottom';
  // the legend the renderer shows anyway stays implicit
  if (out.legend === (out.series.length > 1 || type === 'pie' || type === 'doughnut' ? 'bottom' : 'none')) delete out.legend;
  const interp = at(plotProps, NS_CHART, 'interpolation');
  if (interp && interp !== 'none') out.smooth = true;
  if (seriesEls.some((s) => {
    const v = at(propsOf(s, 'chart-properties'), NS_CHART, 'data-label-number');
    return v !== undefined && v !== 'none';
  }))
    out.dataLabels = true;
  if (type !== 'pie' && type !== 'doughnut' && yAxis && !kid(yAxis, NS_CHART, 'grid')) out.gridlines = false;
  const pt = (el: Element | undefined) => {
    const v = at(el ? kid(styles.get(at(el, NS_CHART, 'style-name') ?? ''), NS_STYLE, 'text-properties') : undefined, NS_FO, 'font-size');
    const m = v ? /^([\d.]+)pt$/.exec(v) : null;
    return m ? Math.round(Number(m[1]) * 10) / 10 : undefined;
  };
  const size = pt(xAxis) ?? pt(yAxis) ?? pt(legend);
  if (size && Math.abs(size - 14) > 0.05) out.fontSize = size;
  const tsize = pt(kid(chartEl, NS_CHART, 'title'));
  if (tsize && out.title && Math.abs(tsize - Math.round((size ?? 14) * 1.29 * 10) / 10) > 0.05) out.titleSize = tsize;
  return out;
}
