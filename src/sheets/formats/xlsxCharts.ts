import { strFromU8, strToU8, unzipSync, zipSync, type Unzipped } from 'fflate';
import { parseFormula } from '../engine/parser';
import type { ChartSpec, ChartType } from '../model/types';
import { colName } from '../model/address';
import { PALETTES } from '../charts/chartConfig';

/* =============================================================== reading */

function attr(xml: string, tag: string, name: string): string | null {
  const m = new RegExp(`<${tag}\\b[^>]*\\b${name}="([^"]*)"`).exec(xml);
  return m ? m[1] : null;
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

function rels(zip: Unzipped, path: string): Map<string, string> {
  const dir = path.split('/').slice(0, -1).join('/');
  const file = `${dir}/_rels/${path.split('/').pop()}.rels`;
  const out = new Map<string, string>();
  const data = zip[file];
  if (!data) return out;
  const xml = strFromU8(data);
  for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], 'Relationship', 'Id');
    const target = attr(m[0], 'Relationship', 'Target');
    if (id && target) out.set(id, resolvePath(path, target));
  }
  return out;
}

const TYPE_MAP: Record<string, ChartType> = {
  barChart: 'column',
  bar3DChart: 'column',
  lineChart: 'line',
  line3DChart: 'line',
  pieChart: 'pie',
  pie3DChart: 'pie',
  doughnutChart: 'doughnut',
  areaChart: 'area',
  area3DChart: 'area',
  scatterChart: 'scatter',
  radarChart: 'radar',
  bubbleChart: 'bubble',
};

interface RefBox {
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function refBox(f: string): RefBox | null {
  try {
    const n = parseFormula(f);
    if (n.t === 'ref' && n.sheet) return { sheet: n.sheet, r1: n.r, c1: n.c, r2: n.r, c2: n.c };
    if (n.t === 'range' && n.sheet) return { sheet: n.sheet, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 };
  } catch {
    /* ignore */
  }
  return null;
}

function chartFromXml(xml: string): Omit<ChartSpec, 'anchor' | 'w' | 'h' | 'id'> | null {
  const plot = /<c:plotArea>([\s\S]*?)<\/c:plotArea>/.exec(xml)?.[1] ?? xml;
  const kinds = [...plot.matchAll(/<c:(\w+Chart)>/g)].map((m) => m[1]);
  if (!kinds.length) return null;
  let type: ChartType = TYPE_MAP[kinds[0]] ?? 'column';
  if (kinds.length > 1 && kinds.includes('barChart') && kinds.includes('lineChart')) type = 'combo';
  if (type === 'column' && /<c:barDir val="bar"\/>/.test(plot)) type = 'bar';
  const grouping = /<c:grouping val="(\w+)"\/>/.exec(plot)?.[1];
  const boxes: RefBox[] = [];
  for (const m of plot.matchAll(/<c:f>([^<]+)<\/c:f>/g)) {
    const b = refBox(m[1].replace(/&apos;/g, "'").replace(/&amp;/g, '&'));
    if (b) boxes.push(b);
  }
  if (!boxes.length) return null;
  const sheet = boxes[0].sheet;
  const same = boxes.filter((b) => b.sheet === sheet);
  const box = { r1: Math.min(...same.map((b) => b.r1)), c1: Math.min(...same.map((b) => b.c1)), r2: Math.max(...same.map((b) => b.r2)), c2: Math.max(...same.map((b) => b.c2)) };
  // series stored along rows when each value ref is a single row
  const valRefs = [...plot.matchAll(/<c:val>[\s\S]*?<c:f>([^<]+)<\/c:f>/g)].map((m) => refBox(m[1]));
  const seriesInRows = valRefs.length > 0 && valRefs.every((b) => b && b.r1 === b.r2 && b.c2 > b.c1);
  const titleText = [...(/<c:title>([\s\S]*?)<\/c:title>/.exec(xml)?.[1] ?? '').matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('');
  const legendPos = /<c:legendPos val="(\w)"\/>/.exec(xml)?.[1];
  const hasLegend = /<c:legend>/.test(xml);
  const quote = /^[A-Za-z_][\w.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  return {
    type,
    source: `${quote}!$${colName(box.c1)}$${box.r1 + 1}:$${colName(box.c2)}$${box.r2 + 1}`,
    seriesInRows,
    stacked: grouping === 'percentStacked' ? 'percent' : grouping === 'stacked' ? true : undefined,
    title: titleText ? decodeXml(titleText) : undefined,
    legend: !hasLegend ? 'none' : legendPos === 't' ? 'top' : legendPos === 'r' ? 'right' : legendPos === 'l' ? 'left' : 'bottom',
    smooth: /<c:smooth val="1"\/>/.test(plot),
    dataLabels: /<c:showVal val="1"\/>/.test(plot),
    palette: 'office',
  };
}

function decodeXml(s: string): string {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

/** Charts per worksheet index (in workbook order). Sizes use EMU→px and cell anchors. */
export function readXlsxCharts(bytes: Uint8Array): Map<number, ChartSpec[]> {
  const out = new Map<number, ChartSpec[]>();
  let zip: Unzipped;
  try {
    zip = unzipSync(bytes, { filter: (f) => f.name.startsWith('xl/') && (f.name.endsWith('.xml') || f.name.endsWith('.rels')) });
  } catch {
    return out;
  }
  const wbXml = zip['xl/workbook.xml'] ? strFromU8(zip['xl/workbook.xml']) : '';
  const wbRels = rels(zip, 'xl/workbook.xml');
  const sheets = [...wbXml.matchAll(/<sheet\b[^>]*>/g)].map((m) => attr(m[0], 'sheet', 'r:id'));
  sheets.forEach((rid, index) => {
    const sheetPath = rid ? wbRels.get(rid) : undefined;
    if (!sheetPath || !zip[sheetPath]) return;
    const sheetXml = strFromU8(zip[sheetPath]);
    const drawRid = attr(sheetXml, 'drawing', 'r:id');
    if (!drawRid) return;
    const drawPath = rels(zip, sheetPath).get(drawRid);
    if (!drawPath || !zip[drawPath]) return;
    const drawXml = strFromU8(zip[drawPath]);
    const drawRels = rels(zip, drawPath);
    const specs: ChartSpec[] = [];
    for (const m of drawXml.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:\1>/g)) {
      const block = m[0];
      const chartRid = /<c:chart\b[^>]*r:id="([^"]+)"/.exec(block)?.[1];
      if (!chartRid) continue;
      const chartPath = drawRels.get(chartRid);
      if (!chartPath || !zip[chartPath]) continue;
      const spec = chartFromXml(strFromU8(zip[chartPath]));
      if (!spec) continue;
      const num = (tag: string, part: string) => Number(new RegExp(`<xdr:${part}>[\\s\\S]*?<xdr:${tag}>(\\d+)</xdr:${tag}>`).exec(block)?.[1] ?? 0);
      const fromCol = num('col', 'from');
      const fromRow = num('row', 'from');
      const fromColOff = num('colOff', 'from') / 9525;
      const fromRowOff = num('rowOff', 'from') / 9525;
      let w = 480;
      let h = 290;
      const ext = /<xdr:ext cx="(\d+)" cy="(\d+)"/.exec(block) ?? /<a:ext cx="(\d+)" cy="(\d+)"/.exec(block);
      if (m[1] === 'twoCellAnchor') {
        const toCol = num('col', 'to');
        const toRow = num('row', 'to');
        // approximate with default Excel metrics (64 px columns, 20 px rows)
        w = Math.max(120, (toCol - fromCol) * 64 + num('colOff', 'to') / 9525 - fromColOff);
        h = Math.max(80, (toRow - fromRow) * 20 + num('rowOff', 'to') / 9525 - fromRowOff);
        if (ext && Number(ext[1]) > 0) {
          w = Number(ext[1]) / 9525;
          h = Number(ext[2]) / 9525;
        }
      } else if (ext) {
        w = Number(ext[1]) / 9525;
        h = Number(ext[2]) / 9525;
      }
      specs.push({ ...spec, id: `xchart${index}_${specs.length}`, anchor: { r: fromRow, c: fromCol, dx: fromColOff, dy: fromRowOff }, w: Math.round(w), h: Math.round(h) });
    }
    if (specs.length) out.set(index, specs);
  });
  return out;
}

/* =============================================================== writing */

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export interface ChartSeriesRefs {
  name?: string; // formula
  nameText?: string;
  cat?: string; // formula
  val: string; // formula
}

export interface ChartToWrite {
  spec: ChartSpec;
  series: ChartSeriesRefs[];
  /** Number of points per series (pie slices get their own colours). */
  points?: number;
  /** Anchor in cells with pixel offsets, plus the bottom-right cell for twoCellAnchor. */
  from: { c: number; r: number; dx: number; dy: number };
  to: { c: number; r: number; dx: number; dy: number };
}

function chartXml(ch: ChartToWrite, index: number): string {
  const s = ch.spec;
  const pal = PALETTES[s.palette ?? 'affice'] ?? PALETTES.affice;
  const hex = (i: number) => pal[i % pal.length].replace('#', '').toUpperCase();
  const stacked = s.stacked === 'percent' ? 'percentStacked' : s.stacked ? 'stacked' : 'clustered';
  const serXml = (kind: string) =>
    ch.series
      .map((sr, i) => {
        const tx = sr.name ? `<c:tx><c:strRef><c:f>${esc(sr.name)}</c:f></c:strRef></c:tx>` : sr.nameText ? `<c:tx><c:v>${esc(sr.nameText)}</c:v></c:tx>` : '';
        const fill = kind === 'line' || kind === 'scatter' || kind === 'radar' ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${hex(i)}"/></a:solidFill><a:round/></a:ln></c:spPr>` : `<c:spPr><a:solidFill><a:srgbClr val="${hex(i)}"/></a:solidFill></c:spPr>`;
        const pie = kind === 'pie' || kind === 'doughnut';
        const slices = pie ? Array.from({ length: ch.points ?? 0 }, (_, p) => `<c:dPt><c:idx val="${p}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${hex(p)}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('') : '';
        const cat = sr.cat ? (kind === 'scatter' ? `<c:xVal><c:numRef><c:f>${esc(sr.cat)}</c:f></c:numRef></c:xVal>` : `<c:cat><c:strRef><c:f>${esc(sr.cat)}</c:f></c:strRef></c:cat>`) : '';
        const val = kind === 'scatter' ? `<c:yVal><c:numRef><c:f>${esc(sr.val)}</c:f></c:numRef></c:yVal>` : `<c:val><c:numRef><c:f>${esc(sr.val)}</c:f></c:numRef></c:val>`;
        const marker = kind === 'line' ? `<c:marker><c:symbol val="${s.smooth ? 'none' : 'circle'}"/></c:marker>` : kind === 'scatter' ? `<c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>` : '';
        const labels = s.dataLabels ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>` : '';
        const smooth = kind === 'line' || kind === 'scatter' ? `<c:smooth val="${s.smooth ? 1 : 0}"/>` : '';
        return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${pie ? '' : fill}${marker}${slices}${labels}${cat}${val}${smooth}</c:ser>`;
      })
      .join('');
  const axes = `<c:axId val="111${index}"/><c:axId val="222${index}"/>`;
  let plot: string;
  const t = s.type === 'combo' ? 'column' : s.type;
  switch (t) {
    case 'bar':
    case 'column':
      plot = `<c:barChart><c:barDir val="${t === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${stacked}"/><c:varyColors val="0"/>${serXml('bar')}<c:gapWidth val="150"/>${s.stacked ? '<c:overlap val="100"/>' : ''}${axes}</c:barChart>`;
      break;
    case 'line':
      plot = `<c:lineChart><c:grouping val="${s.stacked === 'percent' ? 'percentStacked' : s.stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${serXml('line')}<c:marker val="1"/>${axes}</c:lineChart>`;
      break;
    case 'area':
      plot = `<c:areaChart><c:grouping val="${s.stacked === 'percent' ? 'percentStacked' : s.stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${serXml('area')}${axes}</c:areaChart>`;
      break;
    case 'pie':
    case 'polarArea':
      plot = `<c:pieChart><c:varyColors val="1"/>${serXml('pie')}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case 'doughnut':
      plot = `<c:doughnutChart><c:varyColors val="1"/>${serXml('doughnut')}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
      break;
    case 'scatter':
    case 'bubble':
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${serXml('scatter')}${axes}</c:scatterChart>`;
      break;
    case 'radar':
      plot = `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${serXml('radar')}${axes}</c:radarChart>`;
      break;
    default:
      plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${serXml('bar')}${axes}</c:barChart>`;
  }
  const pieLike = ['pie', 'doughnut', 'polarArea'].includes(t);
  const axisTitle = (text?: string) => (text ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '');
  const horizontal = t === 'bar';
  const catAx =
    t === 'scatter' || t === 'bubble'
      ? `<c:valAx><c:axId val="111${index}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${axisTitle(s.xTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="222${index}"/><c:crosses val="autoZero"/></c:valAx>`
      : `<c:catAx><c:axId val="111${index}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/>${axisTitle(horizontal ? s.yTitle : s.xTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="222${index}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
  const valAx = `<c:valAx><c:axId val="222${index}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/>${s.gridlines !== false ? '<c:majorGridlines/>' : ''}${axisTitle(horizontal ? s.xTitle : s.yTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="111${index}"/><c:crosses val="autoZero"/></c:valAx>`;
  const legend = s.legend === 'none' ? '' : `<c:legend><c:legendPos val="${s.legend === 'top' ? 't' : s.legend === 'right' ? 'r' : s.legend === 'left' ? 'l' : 'b'}"/><c:overlay val="0"/></c:legend>`;
  const title = s.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${esc(s.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  let combo = '';
  if (s.type === 'combo' && ch.series.length > 1) {
    // first series as columns, the rest as lines on the same axes
    const rest: ChartToWrite = { ...ch, series: ch.series.slice(1) };
    const lineSer = chartSeriesXml(rest, pal, 1);
    plot = plot.replace(/<c:ser>[\s\S]*<\/c:ser>/, (m) => m.split('</c:ser>')[0] + '</c:ser>');
    combo = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${lineSer}<c:marker val="1"/>${axes}</c:lineChart>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${combo}${pieLike ? '' : catAx + valAx}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>`;
}

function chartSeriesXml(ch: ChartToWrite, pal: string[], offset: number): string {
  return ch.series
    .map((sr, k) => {
      const i = k + offset;
      const hex = pal[i % pal.length].replace('#', '').toUpperCase();
      const tx = sr.name ? `<c:tx><c:strRef><c:f>${esc(sr.name)}</c:f></c:strRef></c:tx>` : '';
      const cat = sr.cat ? `<c:cat><c:strRef><c:f>${esc(sr.cat)}</c:f></c:strRef></c:cat>` : '';
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="circle"/></c:marker>${cat}<c:val><c:numRef><c:f>${esc(sr.val)}</c:f></c:numRef></c:val><c:smooth val="0"/></c:ser>`;
    })
    .join('');
}

const EMU = 9525;

function anchorXml(ch: ChartToWrite, rid: string, id: number): string {
  const { from, to } = ch;
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${from.c}</xdr:col><xdr:colOff>${Math.round(from.dx * EMU)}</xdr:colOff><xdr:row>${from.r}</xdr:row><xdr:rowOff>${Math.round(from.dy * EMU)}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${to.c}</xdr:col><xdr:colOff>${Math.round(to.dx * EMU)}</xdr:colOff><xdr:row>${to.r}</xdr:row><xdr:rowOff>${Math.round(to.dy * EMU)}</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id + 2}" name="Chart ${id + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(ch.spec.w * EMU)}" cy="${Math.round(ch.spec.h * EMU)}"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

/** Adds chart parts to an .xlsx produced by ExcelJS (which cannot write charts). */
export function injectXlsxCharts(bytes: Uint8Array, perSheet: Map<number, ChartToWrite[]>): Uint8Array {
  if (![...perSheet.values()].some((l) => l.length)) return bytes;
  const zip = unzipSync(bytes);
  const put = (path: string, text: string) => (zip[path] = strToU8(text));
  const get = (path: string) => (zip[path] ? strFromU8(zip[path]) : '');
  let ct = get('[Content_Types].xml');
  const wbRels = rels(zip, 'xl/workbook.xml');
  const wbXml = get('xl/workbook.xml');
  const sheetPaths = [...wbXml.matchAll(/<sheet\b[^>]*>/g)].map((m) => wbRels.get(attr(m[0], 'sheet', 'r:id') ?? '') ?? '');
  let chartNo = 1;
  let drawingNo = Object.keys(zip).filter((k) => /^xl\/drawings\/drawing\d+\.xml$/.test(k)).length + 1;
  for (const [index, charts] of perSheet) {
    if (!charts.length) continue;
    const sheetPath = sheetPaths[index];
    if (!sheetPath || !zip[sheetPath]) continue;
    let sheetXml = get(sheetPath);
    const sheetRelsPath = `${sheetPath.split('/').slice(0, -1).join('/')}/_rels/${sheetPath.split('/').pop()}.rels`;
    let sheetRels = get(sheetRelsPath) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const existingDrawRid = attr(sheetXml, 'drawing', 'r:id');
    let drawPath: string;
    let drawXml: string;
    let drawRelsXml: string;
    if (existingDrawRid) {
      drawPath = rels(zip, sheetPath).get(existingDrawRid)!;
      drawXml = get(drawPath);
      drawRelsXml = get(`xl/drawings/_rels/${drawPath.split('/').pop()}.rels`) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    } else {
      drawPath = `xl/drawings/drawing${drawingNo++}.xml`;
      drawXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></xdr:wsDr>`;
      drawRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
      const rid = `rIdAfficeDraw${index}`;
      sheetRels = sheetRels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawPath.split('/').pop()}"/></Relationships>`);
      // <drawing> must come after pageSetup/headerFooter etc.; before legacyDrawing/tableParts/extLst
      const insertBefore = /<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(sheetXml);
      const tag = `<drawing r:id="${rid}"/>`;
      if (insertBefore) sheetXml = sheetXml.slice(0, insertBefore.index) + tag + sheetXml.slice(insertBefore.index);
      else sheetXml = sheetXml.replace('</worksheet>', `${tag}</worksheet>`);
      if (!/xmlns:r=/.test(sheetXml.slice(0, 500))) sheetXml = sheetXml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
      ct = ct.replace('</Types>', `<Override PartName="/${drawPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
    }
    if (!/xmlns:c=/.test(drawXml)) drawXml = drawXml.replace('<xdr:wsDr ', '<xdr:wsDr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ');
    // drawings made for pictures declare r: only on each picture; the chart anchors need it on the root
    if (!/<xdr:wsDr\b[^>]*\sxmlns:r=/.test(drawXml)) drawXml = drawXml.replace('<xdr:wsDr ', '<xdr:wsDr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    let anchors = '';
    charts.forEach((ch, k) => {
      const chartPath = `xl/charts/chart${chartNo}.xml`;
      put(chartPath, chartXml(ch, chartNo));
      ct = ct.replace('</Types>', `<Override PartName="/${chartPath}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
      const rid = `rIdAfficeChart${chartNo}`;
      drawRelsXml = drawRelsXml.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNo}.xml"/></Relationships>`);
      anchors += anchorXml(ch, rid, 1000 + k);
      chartNo++;
    });
    drawXml = drawXml.replace('</xdr:wsDr>', `${anchors}</xdr:wsDr>`);
    put(drawPath, drawXml);
    put(`xl/drawings/_rels/${drawPath.split('/').pop()}.rels`, drawRelsXml);
    put(sheetPath, sheetXml);
    put(sheetRelsPath, sheetRels);
  }
  put('[Content_Types].xml', ct);
  return zipSync(zip, { level: 6 });
}
