import type { ChartConfiguration, ChartDataset, Plugin } from 'chart.js';
import { cssFontStack } from '@/lib/fonts';
import { resolveColor, type SlideChart, type Theme } from '../model';

const ACCENTS = ['@accent1', '@accent2', '@accent3', '@accent4', '@accent5', '@accent6'];

function alpha(css: string, a: number): string {
  const m = /^#([0-9a-f]{6})$/i.exec(css);
  if (m) {
    const n = parseInt(m[1], 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
  }
  return css;
}

export function seriesColor(chart: SlideChart, i: number, theme: Theme): string {
  const c = chart.series[i]?.color;
  if (c) return resolveColor(c, theme);
  const base = resolveColor(ACCENTS[i % 6], theme);
  return i < 6 ? base : alpha(base, 0.6);
}

const labelsPlugin: Plugin = {
  id: 'slideLabels',
  afterDatasetsDraw(chart) {
    const opts = (chart.options.plugins as Record<string, { show?: boolean; color?: string; font?: string }>).slideLabels;
    if (!opts?.show) return;
    const ctx = chart.ctx;
    ctx.save();
    ctx.fillStyle = opts.color ?? '#333';
    ctx.font = `600 15px ${opts.font ?? 'sans-serif'}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    chart.data.datasets.forEach((ds, i) => {
      const meta = chart.getDatasetMeta(i);
      if (meta.hidden) return;
      meta.data.forEach((el, j) => {
        const v = (ds.data as (number | null)[])[j];
        if (v === null || v === undefined) return;
        const pos = el.tooltipPosition(true);
        ctx.fillText(String(Math.round(Number(v) * 100) / 100), pos.x ?? 0, (pos.y ?? 0) - 4);
      });
    });
    ctx.restore();
  },
};

export function slideChartConfig(chart: SlideChart, theme: Theme, animate: boolean, fontScale = 1): ChartConfiguration {
  const text = resolveColor('@tx1', theme);
  const grid = alpha(resolveColor('@tx1', theme).startsWith('#') ? resolveColor('@tx1', theme) : '#000000', 0.12);
  const font = cssFontStack(theme.fonts.minor);
  const pie = chart.type === 'pie' || chart.type === 'doughnut';
  const stacked = !!chart.stacked;
  let type: ChartConfiguration['type'] = 'bar';
  let datasets: ChartDataset[];
  switch (chart.type) {
    case 'line':
    case 'area':
      type = 'line';
      datasets = chart.series.map((s, i) => {
        const c = seriesColor(chart, i, theme);
        return { label: s.name, data: s.values, borderColor: c, backgroundColor: chart.type === 'area' ? alpha(c, 0.35) : c, fill: chart.type === 'area' ? (stacked && i > 0 ? '-1' : 'origin') : false, tension: chart.smooth ? 0.35 : 0, pointRadius: 3, borderWidth: 3, spanGaps: true };
      });
      break;
    case 'pie':
    case 'doughnut': {
      type = chart.type;
      const s = chart.series[0] ?? { name: '', values: [] };
      datasets = [{ label: s.name, data: s.values.map((v) => v ?? 0), backgroundColor: chart.categories.map((_, i) => seriesColor({ ...chart, series: [] }, i, theme)), borderColor: resolveColor('@bg1', theme), borderWidth: 2 }];
      break;
    }
    case 'scatter':
      type = 'scatter';
      datasets = chart.series.map((s, i) => ({ label: s.name, data: s.values.map((y, k) => ({ x: Number(chart.categories[k]) || k + 1, y })).filter((p) => p.y !== null) as { x: number; y: number }[], backgroundColor: seriesColor(chart, i, theme), pointRadius: 5 }));
      break;
    case 'radar':
      type = 'radar';
      datasets = chart.series.map((s, i) => {
        const c = seriesColor(chart, i, theme);
        return { label: s.name, data: s.values, borderColor: c, backgroundColor: alpha(c, 0.2), pointRadius: 3 };
      });
      break;
    default:
      type = 'bar';
      datasets = chart.series.map((s, i) => ({ label: s.name, data: s.values, backgroundColor: seriesColor(chart, i, theme), borderRadius: 3, maxBarThickness: 80 }));
  }
  if (chart.stacked === 'percent' && !pie) {
    const n = chart.categories.length;
    const totals = Array.from({ length: n }, (_, k) => datasets.reduce((acc, ds) => acc + Math.abs(Number((ds.data as (number | null)[])[k] ?? 0)), 0));
    for (const ds of datasets) ds.data = (ds.data as (number | null)[]).map((v, k) => (v === null ? null : totals[k] ? (100 * v) / totals[k] : 0));
  }
  const horizontal = chart.type === 'bar';
  const fs = (n: number) => Math.round(n * fontScale);
  const legend = chart.legend ?? (chart.series.length > 1 || pie ? 'bottom' : 'none');
  const catAxis = { stacked, grid: { display: false }, ticks: { color: text, font: { family: font, size: fs(16) } } };
  const valAxis = { stacked, beginAtZero: true, grid: { color: grid, display: chart.gridlines !== false }, ticks: { color: text, font: { family: font, size: fs(15) } }, ...(chart.stacked === 'percent' ? { max: 100 } : {}) };
  return {
    type,
    data: { labels: chart.type === 'scatter' ? undefined : chart.categories, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: animate ? { duration: 500 } : false,
      indexAxis: horizontal ? 'y' : 'x',
      layout: { padding: 6 },
      plugins: {
        title: { display: !!chart.title, text: chart.title ?? '', color: text, font: { family: font, size: fs(24), weight: 600 } },
        legend: { display: legend !== 'none', position: legend === 'none' ? 'bottom' : legend, labels: { color: text, usePointStyle: true, boxWidth: fs(10), boxHeight: fs(10), padding: fs(18), font: { family: font, size: fs(16) } } },
        tooltip: { enabled: animate },
        slideLabels: { show: !!chart.dataLabels, color: text, font },
      } as Record<string, unknown>,
      scales: (pie
        ? {}
        : chart.type === 'radar'
          ? { r: { grid: { color: grid }, angleLines: { color: grid }, pointLabels: { color: text, font: { family: font, size: fs(15) } }, ticks: { color: text, backdropColor: 'transparent' } } }
          : { x: horizontal ? valAxis : catAxis, y: horizontal ? catAxis : valAxis }) as never,
    },
    plugins: [labelsPlugin],
  };
}

export const SAMPLE_CHART: SlideChart = {
  type: 'column',
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', values: [2, 2, 3, 5] },
  ],
  legend: 'bottom',
};
