import type { ChartConfiguration, ChartDataset, Plugin } from 'chart.js';
import { parseFormula } from '../engine/parser';
import { isErr } from '../engine/values';
import { formatValue, isDateFormat } from '../format/numfmt';
import type { ChartSpec } from '../model/types';
import type { Sheet } from '../model/workbook';
import type { SheetDoc } from '../doc';

export const PALETTES: Record<string, string[]> = {
  affice: ['#2f6dff', '#17a35a', '#f26a26', '#8e4ec6', '#e5484d', '#0090ff', '#f5a524', '#12a594', '#d6409f', '#6e56cf'],
  office: ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e', '#636363', '#997300'],
  pastel: ['#8fb3ff', '#8ee0b2', '#ffb38a', '#c9a7f5', '#ff9fa2', '#8fd3ff', '#ffd68a', '#9be3da', '#f5a3cf', '#b4a7f5'],
  vivid: ['#ff3b30', '#ff9500', '#ffcc00', '#34c759', '#00c7be', '#30b0c7', '#007aff', '#5856d6', '#af52de', '#ff2d55'],
  mono: ['#1f3b73', '#2f5aa8', '#4a7bd0', '#7aa1e0', '#a9c3ec', '#d2e0f6', '#0f2247', '#3c6ebf', '#94b6ea', '#c3d6f3'],
  earth: ['#8c6d46', '#b5a27d', '#6b8e5a', '#c28b5f', '#7a6a5b', '#a8b58a', '#5d4a36', '#d0b38a', '#4f6b44', '#9d7b5c'],
};

export interface SeriesData {
  categories: string[];
  series: { name: string; values: (number | null)[] }[];
  xNumeric: boolean;
}

/** Resolves a chart's source reference to a sheet + range. */
export function resolveSource(doc: SheetDoc, spec: Pick<ChartSpec, 'source'>, home: Sheet): { sheet: Sheet; r1: number; c1: number; r2: number; c2: number } | null {
  try {
    const node = parseFormula(spec.source.replace(/^=/, ''));
    if (node.t !== 'range' && node.t !== 'ref') return null;
    const sheet = node.sheet ? doc.wb.sheetByName(node.sheet) : home;
    if (!sheet) return null;
    if (node.t === 'ref') return { sheet, r1: node.r, c1: node.c, r2: node.r, c2: node.c };
    const ext = doc.engine.extent(sheet.id);
    return { sheet, r1: node.r1, c1: node.c1, r2: Math.min(node.r2, Math.max(node.r1, ext.rows)), c2: Math.min(node.c2, Math.max(node.c1, ext.cols)) };
  } catch {
    return null;
  }
}

export function chartData(doc: SheetDoc, spec: ChartSpec, home: Sheet): SeriesData {
  const src = resolveSource(doc, spec, home);
  if (!src) return { categories: [], series: [], xNumeric: false };
  const { sheet } = src;
  const grid: unknown[][] = [];
  for (let r = src.r1; r <= src.r2; r++) {
    const row: unknown[] = [];
    for (let c = src.c1; c <= src.c2; c++) row.push(doc.value(sheet, r, c));
    grid.push(row);
  }
  const rows = grid.length;
  const cols = grid[0]?.length ?? 0;
  const byRows = spec.seriesInRows ?? cols > rows;
  // orient so that series are columns
  const g = byRows ? transpose(grid) : grid;
  const fmtCell = (i: number, j: number) => {
    const r = byRows ? src.r1 + j : src.r1 + i;
    const c = byRows ? src.c1 + i : src.c1 + j;
    const v = g[i][j];
    if (v === null || v === undefined) return '';
    const fmt = doc.styleOf(sheet, r, c).numFmt;
    return formatValue(v, fmt).text;
  };
  const R = g.length;
  const C = g[0]?.length ?? 0;
  const isNum = (v: unknown) => typeof v === 'number';
  const headerRow = spec.headerRow ?? (R > 1 && g[0].slice(spec.type === 'scatter' ? 0 : 1).some((v) => typeof v === 'string') && g.slice(1).some((row) => row.some(isNum)));
  const firstCol = g.slice(headerRow ? 1 : 0).map((row) => row[0]);
  const firstColDateFmt = (() => {
    const r = byRows ? src.r1 : src.r1 + (headerRow ? 1 : 0);
    const c = byRows ? src.c1 + (headerRow ? 1 : 0) : src.c1;
    return isDateFormat(doc.styleOf(sheet, r, c).numFmt);
  })();
  const headerCol = spec.headerCol ?? (C > 1 && (firstCol.some((v) => typeof v === 'string') || firstColDateFmt || (spec.type === 'scatter' && firstCol.every((v) => isNum(v) || v === null))));
  const start = headerRow ? 1 : 0;
  const categories = headerCol ? g.slice(start).map((_, i) => fmtCell(i + start, 0)) : g.slice(start).map((_, i) => String(i + 1));
  const series: SeriesData['series'] = [];
  for (let j = headerCol ? 1 : 0; j < C; j++) {
    const name = headerRow ? fmtCell(0, j) || `Series ${series.length + 1}` : `Series ${series.length + 1}`;
    const values = g.slice(start).map((row) => {
      const v = row[j];
      return typeof v === 'number' && !isErr(v) ? v : null;
    });
    series.push({ name, values });
  }
  const xNumeric = headerCol && firstCol.every((v) => isNum(v) || v === null) && !firstColDateFmt;
  return { categories, series, xNumeric };
}

function transpose<T>(g: T[][]): T[][] {
  const R = g.length;
  const C = g[0]?.length ?? 0;
  const out: T[][] = [];
  for (let c = 0; c < C; c++) {
    const row: T[] = [];
    for (let r = 0; r < R; r++) row.push(g[r][c]);
    out.push(row);
  }
  return out;
}

const alpha = (hex: string, a: number) => hex + Math.round(a * 255).toString(16).padStart(2, '0');

/** Draws values above bars/points when data labels are on. */
const dataLabels: Plugin = {
  id: 'afficeLabels',
  afterDatasetsDraw(chart, _args, opts: { show?: boolean; color?: string }) {
    if (!opts?.show) return;
    const { ctx } = chart;
    ctx.save();
    ctx.font = '600 11px Inter, system-ui, sans-serif';
    ctx.fillStyle = opts.color ?? '#333';
    ctx.textAlign = 'center';
    chart.data.datasets.forEach((ds, i) => {
      const meta = chart.getDatasetMeta(i);
      if (meta.hidden) return;
      meta.data.forEach((el, j) => {
        const v = ds.data[j] as number | { y: number } | null;
        const n = typeof v === 'number' ? v : v && typeof v === 'object' ? v.y : null;
        if (n === null || n === undefined) return;
        const p = el.tooltipPosition(true);
        const label = Math.abs(n) >= 1000 ? n.toLocaleString(undefined, { maximumFractionDigits: 1 }) : String(Math.round(n * 100) / 100);
        ctx.fillText(label, p.x ?? 0, (p.y ?? 0) - 6);
      });
    });
    ctx.restore();
  },
};

export function buildChartConfig(doc: SheetDoc, spec: ChartSpec, home: Sheet, dark: boolean, animate: boolean): ChartConfiguration {
  const data = chartData(doc, spec, home);
  const pal = PALETTES[spec.palette ?? 'affice'] ?? PALETTES.affice;
  const text = dark ? '#d5d8e0' : '#3a4050';
  const grid = dark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.07)';
  const stacked = !!spec.stacked;
  const percent = spec.stacked === 'percent';
  const pie = spec.type === 'pie' || spec.type === 'doughnut' || spec.type === 'polarArea';
  let type: ChartConfiguration['type'] = 'bar';
  let datasets: ChartDataset[] = [];
  switch (spec.type) {
    case 'column':
    case 'bar':
      type = 'bar';
      datasets = data.series.map((s, i) => ({ label: s.name, data: s.values, backgroundColor: pal[i % pal.length], borderRadius: 3, maxBarThickness: 64 }));
      break;
    case 'line':
    case 'area':
      type = 'line';
      datasets = data.series.map((s, i) => ({
        label: s.name,
        data: s.values,
        borderColor: pal[i % pal.length],
        backgroundColor: spec.type === 'area' ? alpha(pal[i % pal.length], stacked ? 0.75 : 0.25) : pal[i % pal.length],
        fill: spec.type === 'area' ? (stacked ? (i === 0 ? 'origin' : '-1') : 'origin') : false,
        tension: spec.smooth ? 0.35 : 0,
        pointRadius: data.categories.length > 60 ? 0 : 3,
        borderWidth: 2,
        spanGaps: true,
      }));
      break;
    case 'pie':
    case 'doughnut':
    case 'polarArea': {
      type = spec.type;
      const s = data.series[0] ?? { name: '', values: [] };
      datasets = [{ label: s.name, data: s.values.map((v) => v ?? 0), backgroundColor: data.categories.map((_, i) => (spec.type === 'polarArea' ? alpha(pal[i % pal.length], 0.7) : pal[i % pal.length])), borderColor: dark ? '#1f2128' : '#fff', borderWidth: 2 }];
      break;
    }
    case 'scatter':
    case 'bubble': {
      type = spec.type;
      const xs = data.xNumeric ? data.categories.map(Number) : data.categories.map((_, i) => i + 1);
      if (spec.type === 'bubble') {
        const ys = data.series[0]?.values ?? [];
        const rs = data.series[1]?.values ?? [];
        const maxR = Math.max(1, ...rs.map((v) => Math.abs(v ?? 0)));
        datasets = [{ label: data.series[0]?.name ?? 'Series 1', data: xs.map((x, i) => ({ x, y: ys[i] ?? 0, r: 4 + (Math.abs(rs[i] ?? 0) / maxR) * 20 })), backgroundColor: alpha(pal[0], 0.6), borderColor: pal[0] }];
      } else {
        datasets = data.series.map((s, i) => ({
          label: s.name,
          data: xs.map((x, k) => ({ x, y: s.values[k] })).filter((p) => p.y !== null) as { x: number; y: number }[],
          backgroundColor: pal[i % pal.length],
          borderColor: pal[i % pal.length],
          pointRadius: 4,
          showLine: !!spec.smooth,
        }));
      }
      break;
    }
    case 'radar':
      type = 'radar';
      datasets = data.series.map((s, i) => ({ label: s.name, data: s.values, borderColor: pal[i % pal.length], backgroundColor: alpha(pal[i % pal.length], 0.2), pointRadius: 3 }));
      break;
    case 'combo':
      type = 'bar';
      datasets = data.series.map((s, i) =>
        i === 0
          ? ({ type: 'bar', label: s.name, data: s.values, backgroundColor: pal[0], borderRadius: 3, order: 2 } as ChartDataset)
          : ({ type: 'line', label: s.name, data: s.values, borderColor: pal[i % pal.length], backgroundColor: pal[i % pal.length], tension: spec.smooth ? 0.35 : 0, order: 1, yAxisID: i === 1 && data.series.length === 2 ? 'y2' : 'y' } as ChartDataset),
      );
      break;
  }
  const legend = spec.legend ?? (data.series.length > 1 || pie ? 'bottom' : 'none');
  const horizontal = spec.type === 'bar';
  const axisTitle = (t?: string) => (t ? { display: true, text: t, color: text, font: { size: 12, weight: 600 } } : { display: false });
  const xy = spec.type === 'scatter' || spec.type === 'bubble';
  // category axis and value axis; a horizontal bar chart swaps which one is x
  const catAxis: Record<string, unknown> = {
    stacked: stacked && spec.type !== 'line',
    grid: { color: grid, display: false },
    ticks: { color: text, maxRotation: 45, autoSkip: true },
    title: axisTitle(spec.xTitle),
    ...(xy ? { type: 'linear', grid: { color: grid, display: spec.gridlines !== false } } : {}),
  };
  const valAxis: Record<string, unknown> = {
    stacked,
    grid: { color: grid, display: spec.gridlines !== false },
    ticks: { color: text },
    beginAtZero: !xy,
    title: axisTitle(spec.yTitle),
    ...(xy ? { type: 'linear' } : {}),
    ...(percent ? { max: 100 } : {}),
  };
  const scales: Record<string, unknown> = pie || spec.type === 'radar'
    ? spec.type === 'radar'
      ? { r: { grid: { color: grid }, angleLines: { color: grid }, pointLabels: { color: text }, ticks: { color: text, backdropColor: 'transparent' } } }
      : {}
    : {
        x: horizontal ? valAxis : catAxis,
        y: horizontal ? catAxis : valAxis,
        ...(spec.type === 'combo' && data.series.length === 2 ? { y2: { position: 'right', grid: { display: false }, ticks: { color: text } } } : {}),
      };
  if (percent && !pie) {
    // normalise each category to 100%
    const n = data.categories.length;
    const totals = Array.from({ length: n }, (_, k) => datasets.reduce((s, ds) => s + Math.abs(Number((ds.data as (number | null)[])[k] ?? 0)), 0));
    for (const ds of datasets) ds.data = (ds.data as (number | null)[]).map((v, k) => (v === null ? null : totals[k] ? (100 * v) / totals[k] : 0));
  }
  return {
    type,
    data: { labels: spec.type === 'scatter' || spec.type === 'bubble' ? undefined : data.categories, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: animate ? { duration: 350 } : false,
      indexAxis: horizontal ? 'y' : 'x',
      layout: { padding: 8 },
      plugins: {
        title: { display: !!spec.title, text: spec.title ?? '', color: text, font: { size: 15, weight: 600 }, padding: { bottom: 8 } },
        legend: { display: legend !== 'none', position: legend === 'none' ? 'bottom' : legend, labels: { color: text, usePointStyle: true, pointStyle: spec.type === 'line' || spec.type === 'scatter' || spec.type === 'radar' ? 'circle' : 'rectRounded', boxWidth: 8, boxHeight: 8, padding: 14 } },
        tooltip: { enabled: true },
        afficeLabels: { show: !!spec.dataLabels, color: text },
      } as Record<string, unknown>,
      scales: scales as never,
    },
    plugins: [dataLabels],
  } as ChartConfiguration;
}
