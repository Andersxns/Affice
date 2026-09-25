/** DrawingML charts for .pptx: writing (with cached values and an embedded workbook) and reading. */
import { escapeXml } from '@/lib/utils';
import type { SlideChart, SlideChartType } from '../model';

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

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

export type ColorXml = (color: string) => string;

function seriesXml(chart: SlideChart, i: number, colorXml: ColorXml, kind: 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'radar'): string {
  const s = chart.series[i];
  const n = chart.categories.length;
  const col = colLetter(i + 1);
  const color = s.color ?? `@accent${(i % 6) + 1}`;
  const tx = `<c:tx><c:strRef><c:f>Sheet1!$${col}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
  const cat = `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f><c:strCache><c:ptCount val="${n}"/>${chart.categories.map((c, k) => `<c:pt idx="${k}"><c:v>${escapeXml(c)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>`;
  const nums = `<c:numRef><c:f>Sheet1!$${col}$2:$${col}$${n + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${s.values.map((v, k) => (v === null || v === undefined ? '' : `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numCache></c:numRef>`;
  let sp: string;
  if (kind === 'line' || kind === 'radar') sp = `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>${colorXml(color)}</a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>`;
  else if (kind === 'scatter') sp = `<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill>${colorXml(color)}</a:solidFill></c:spPr></c:marker>`;
  else if (kind === 'pie') sp = chart.categories.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${colorXml(`@accent${(k % 6) + 1}`)}</a:solidFill><a:ln w="19050"><a:solidFill>${colorXml('@bg1')}</a:solidFill></a:ln></c:spPr></c:dPt>`).join('');
  else sp = `<c:spPr><a:solidFill>${colorXml(color)}</a:solidFill></c:spPr>`;
  const inv = kind === 'bar' ? '<c:invertIfNegative val="0"/>' : '';
  const labels = chart.dataLabels ? '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : '';
  if (kind === 'pie') return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${labels}${cat}<c:val>${nums}</c:val></c:ser>`;
  if (kind === 'scatter') {
    const xs = `<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${chart.categories.map((c, k) => `<c:pt idx="${k}"><c:v>${Number(c) || k + 1}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:xVal>`;
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${labels}${xs}<c:yVal>${nums}</c:yVal><c:smooth val="${chart.smooth ? 1 : 0}"/></c:ser>`;
  }
  const smooth = kind === 'line' ? `<c:smooth val="${chart.smooth ? 1 : 0}"/>` : '';
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${inv}${labels}${cat}<c:val>${nums}</c:val>${smooth}</c:ser>`;
}

/** Axis and legend text size in hundredths of a point. */
const textSz = (chart: SlideChart) => Math.round((chart.fontSize ?? 14) * 100);
const titleSz = (chart: SlideChart) => Math.round((chart.titleSize ?? (chart.fontSize ?? 14) * 1.29) * 100);

const AX_CAT = 111111111;
const AX_VAL = 222222222;

function axes(chart: SlideChart, colorXml: ColorXml, horizontal: boolean, scatter: boolean): string {
  const txt = `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${textSz(chart)}"><a:solidFill>${colorXml('@tx1+35')}</a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
  const grid = chart.gridlines === false ? '' : `<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill>${colorXml('@tx1+85')}</a:solidFill></a:ln></c:spPr></c:majorGridlines>`;
  const catPos = horizontal ? 'l' : 'b';
  const valPos = horizontal ? 'b' : 'l';
  const catAx = scatter
    ? `<c:valAx><c:axId val="${AX_CAT}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:solidFill>${colorXml('@tx1+75')}</a:solidFill></a:ln></c:spPr>${txt}<c:crossAx val="${AX_VAL}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
    : `<c:catAx><c:axId val="${AX_CAT}"/><c:scaling><c:orientation val="${horizontal ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:solidFill>${colorXml('@tx1+75')}</a:solidFill></a:ln></c:spPr>${txt}<c:crossAx val="${AX_VAL}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const valAx = `<c:valAx><c:axId val="${AX_VAL}"/><c:scaling><c:orientation val="minMax"/>${chart.stacked === 'percent' ? '<c:max val="1"/>' : ''}</c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/>${grid}<c:numFmt formatCode="${chart.stacked === 'percent' ? '0%' : 'General'}" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txt}<c:crossAx val="${AX_CAT}"/><c:crosses val="${horizontal ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
  return catAx + valAx;
}

/** chartSpace XML for a slide chart. `externalRid` links the embedded workbook. */
export function chartSpaceXml(chart: SlideChart, colorXml: ColorXml, externalRid?: string): string {
  const grouping = chart.stacked === 'percent' ? 'percentStacked' : chart.stacked ? 'stacked' : 'clustered';
  const idxs = chart.series.map((_, i) => i);
  let plot: string;
  switch (chart.type) {
    case 'line':
      plot = `<c:lineChart><c:grouping val="${chart.stacked ? (chart.stacked === 'percent' ? 'percentStacked' : 'stacked') : 'standard'}"/><c:varyColors val="0"/>${idxs.map((i) => seriesXml(chart, i, colorXml, 'line')).join('')}<c:marker val="1"/><c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:lineChart>${axes(chart, colorXml, false, false)}`;
      break;
    case 'area':
      plot = `<c:areaChart><c:grouping val="${chart.stacked ? (chart.stacked === 'percent' ? 'percentStacked' : 'stacked') : 'standard'}"/><c:varyColors val="0"/>${idxs.map((i) => seriesXml(chart, i, colorXml, 'area')).join('')}<c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:areaChart>${axes(chart, colorXml, false, false)}`;
      break;
    case 'pie':
      plot = `<c:pieChart><c:varyColors val="1"/>${seriesXml(chart, 0, colorXml, 'pie')}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case 'doughnut':
      plot = `<c:doughnutChart><c:varyColors val="1"/>${seriesXml(chart, 0, colorXml, 'pie')}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
      break;
    case 'scatter':
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${idxs.map((i) => seriesXml(chart, i, colorXml, 'scatter')).join('')}<c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:scatterChart>${axes(chart, colorXml, false, true)}`;
      break;
    case 'radar':
      plot = `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${idxs.map((i) => seriesXml(chart, i, colorXml, 'radar')).join('')}<c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:radarChart>${axes(chart, colorXml, false, false)}`;
      break;
    default: {
      const horizontal = chart.type === 'bar';
      plot = `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${idxs.map((i) => seriesXml(chart, i, colorXml, 'bar')).join('')}<c:gapWidth val="80"/>${chart.stacked ? '<c:overlap val="100"/>' : '<c:overlap val="-10"/>'}<c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:barChart>${axes(chart, colorXml, horizontal, false)}`;
    }
  }
  const legendPos = chart.legend && chart.legend !== 'none' ? { bottom: 'b', right: 'r', top: 't', left: 'l' }[chart.legend] : chart.legend === 'none' ? null : chart.series.length > 1 || chart.type === 'pie' || chart.type === 'doughnut' ? 'b' : null;
  const legend = legendPos ? `<c:legend><c:legendPos val="${legendPos}"/><c:overlay val="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${textSz(chart)}"><a:solidFill>${colorXml('@tx1+25')}</a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:legend>` : '';
  const title = chart.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${titleSz(chart)}" b="1"><a:solidFill>${colorXml('@tx1')}</a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="en-US" sz="${titleSz(chart)}" b="1"><a:solidFill>${colorXml('@tx1')}</a:solidFill></a:rPr><a:t>${escapeXml(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>${externalRid ? `<c:externalData r:id="${externalRid}"><c:autoUpdate val="0"/></c:externalData>` : ''}</c:chartSpace>`;
}

/** Embedded workbook holding the chart data (so "Edit data" works in PowerPoint). */
export async function chartWorkbook(chart: SlideChart): Promise<Uint8Array> {
  const ExcelJS = (await import('exceljs')).default;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getRow(1).values = ['', ...chart.series.map((s) => s.name)];
  chart.categories.forEach((c, k) => {
    ws.getRow(k + 2).values = [c, ...chart.series.map((s) => s.values[k] ?? null)];
  });
  const buf = await wb.xlsx.writeBuffer();
  return new Uint8Array(buf as ArrayBuffer);
}

/* =============================================================== reading */

const TYPES: Record<string, SlideChartType> = {
  barChart: 'column',
  bar3DChart: 'column',
  lineChart: 'line',
  line3DChart: 'line',
  areaChart: 'area',
  area3DChart: 'area',
  pieChart: 'pie',
  pie3DChart: 'pie',
  ofPieChart: 'pie',
  doughnutChart: 'doughnut',
  scatterChart: 'scatter',
  bubbleChart: 'scatter',
  radarChart: 'radar',
};

function kids(el: Element | null | undefined, name: string): Element[] {
  if (!el) return [];
  return Array.from(el.children).filter((c) => c.localName === name);
}

function kid(el: Element | null | undefined, name: string): Element | null {
  return kids(el, name)[0] ?? null;
}

function cacheValues(ref: Element | null): (string | null)[] {
  if (!ref) return [];
  const cache = kid(ref, 'strCache') ?? kid(ref, 'numCache');
  const lit = ref.localName === 'strLit' || ref.localName === 'numLit' ? ref : null;
  const src = cache ?? lit;
  if (!src) return [];
  const n = Number(kid(src, 'ptCount')?.getAttribute('val') ?? 0);
  const out: (string | null)[] = Array.from({ length: n }, () => null);
  for (const pt of kids(src, 'pt')) {
    const i = Number(pt.getAttribute('idx') ?? 0);
    out[i] = kid(pt, 'v')?.textContent ?? null;
  }
  return out;
}

function dataOf(el: Element | null): (string | null)[] {
  if (!el) return [];
  const inner = kid(el, 'strRef') ?? kid(el, 'numRef') ?? kid(el, 'strLit') ?? kid(el, 'numLit') ?? kid(el, 'multiLvlStrRef');
  if (inner?.localName === 'multiLvlStrRef') {
    const lvl = kid(kid(inner, 'multiLvlStrCache'), 'lvl');
    return kids(lvl, 'pt').map((p) => kid(p, 'v')?.textContent ?? null);
  }
  return cacheValues(inner);
}

/** Parses a chart part into a slide chart (type, categories, series with cached values). */
export function parseChartXml(xml: string, colorOf: (spPr: Element | null) => string | undefined): SlideChart | null {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const plot = doc.getElementsByTagNameNS(C_NS, 'plotArea')[0];
  if (!plot) return null;
  const group = Array.from(plot.children).find((c) => TYPES[c.localName]);
  if (!group) return null;
  let type = TYPES[group.localName];
  if (type === 'column' && kid(group, 'barDir')?.getAttribute('val') === 'bar') type = 'bar';
  const grouping = kid(group, 'grouping')?.getAttribute('val');
  const series = kids(group, 'ser');
  let categories: string[] = [];
  const out: SlideChart = { type, categories: [], series: [] };
  for (const s of series) {
    const name = dataOf(kid(s, 'tx'))[0] ?? kid(s, 'tx')?.textContent ?? `Series ${out.series.length + 1}`;
    const cats = dataOf(kid(s, 'cat') ?? kid(s, 'xVal'));
    const vals = dataOf(kid(s, 'val') ?? kid(s, 'yVal'));
    if (cats.length > categories.length) categories = cats.map((c, i) => c ?? String(i + 1));
    const color = colorOf(kid(s, 'spPr'));
    out.series.push({ name: name ?? '', values: vals.map((v) => (v === null || v === '' ? null : Number(v))), ...(color ? { color } : {}) });
  }
  if (!categories.length) categories = (out.series[0]?.values ?? []).map((_, i) => String(i + 1));
  out.categories = categories;
  out.series = out.series.map((s) => ({ ...s, values: categories.map((_, i) => s.values[i] ?? null) }));
  if (grouping === 'stacked') out.stacked = true;
  if (grouping === 'percentStacked') out.stacked = 'percent';
  // the chart title (the title element directly inside c:chart, not an axis title)
  const chartEl = doc.getElementsByTagNameNS(C_NS, 'chart')[0];
  const titleEl = kid(chartEl, 'title');
  const autoDeleted = kid(chartEl, 'autoTitleDeleted')?.getAttribute('val');
  if (titleEl) {
    const t = Array.from(titleEl.getElementsByTagNameNS(A_NS, 't')).map((x) => x.textContent ?? '').join('');
    // a title without text is an automatic one: the series name for single-series charts
    out.title = t || (out.series.length === 1 ? out.series[0].name : 'Chart Title');
  } else if (out.series.length === 1 && autoDeleted !== '1' && autoDeleted !== 'true') out.title = out.series[0].name;
  // text sizes: axis labels, then the legend, then the chart-wide default (PowerPoint falls back to 18 pt)
  const szOf = (el: Element | null | undefined) => {
    const r = el ? (Array.from(el.getElementsByTagNameNS(A_NS, 'defRPr')).find((x) => x.getAttribute('sz')) ?? Array.from(el.getElementsByTagNameNS(A_NS, 'rPr')).find((x) => x.getAttribute('sz'))) : undefined;
    return r ? Number(r.getAttribute('sz')) / 100 : undefined;
  };
  const axis = kid(plot, 'catAx') ?? kid(plot, 'valAx') ?? kid(plot, 'dateAx');
  const chartSpace = doc.documentElement;
  const size = szOf(kid(axis, 'txPr')) ?? szOf(kid(kid(chartEl, 'legend'), 'txPr')) ?? szOf(kid(chartSpace, 'txPr')) ?? 18;
  out.fontSize = Math.round(size * 10) / 10;
  const tsz = szOf(titleEl);
  if (tsz) out.titleSize = tsz;
  else if (out.title) out.titleSize = Math.round((szOf(kid(chartSpace, 'txPr')) ?? 18) * 1.2 * 10) / 10;
  const legendPos = doc.getElementsByTagNameNS(C_NS, 'legendPos')[0]?.getAttribute('val');
  out.legend = !doc.getElementsByTagNameNS(C_NS, 'legend').length ? 'none' : legendPos === 'r' ? 'right' : legendPos === 't' ? 'top' : legendPos === 'l' ? 'left' : 'bottom';
  const dl = group.getElementsByTagNameNS(C_NS, 'showVal')[0]?.getAttribute('val');
  if (dl === '1') out.dataLabels = true;
  if (kid(group, 'smooth')?.getAttribute('val') === '1' || series.some((s) => kid(s, 'smooth')?.getAttribute('val') === '1')) out.smooth = true;
  return out;
}
