/**
 * Renders slides outside the editor: PNG pictures, PDF / print jobs and static HTML snapshots.
 * Fonts used by the slides are embedded as data URLs so pictures and exported pages look identical.
 */
import '../slides.css';
import { flushSync } from 'react-dom';
import { createRoot } from 'react-dom/client';
import { zipSync } from 'fflate';
import type { PdfJob } from '@/shared/types';
import { BUNDLED_FACES } from '@/lib/fonts.generated';
import { walkEls, type Presentation } from '../model';
import { SlideView } from '../render/SlideView';
import slidesCss from '../slides.css?inline';

/* ------------------------------------------------------------------ fonts */

function usedChars(pres: Presentation): Set<number> {
  const set = new Set<number>();
  const add = (s: string) => {
    for (const ch of s) set.add(ch.codePointAt(0)!);
  };
  for (const s of pres.slides)
    walkEls(s.elements, (e) => {
      if (e.type === 'shape' && e.text) e.text.paras.forEach((p) => p.runs.forEach((r) => add(r.text)));
      if (e.type === 'table') e.rows.forEach((r) => r.cells.forEach((c) => c.text.paras.forEach((p) => p.runs.forEach((x) => add(x.text)))));
      if (e.type === 'chart') [e.chart.title ?? '', ...e.chart.categories, ...e.chart.series.map((x) => x.name)].forEach(add);
    });
  add('0123456789 .,:;!?%-—–•…’“”');
  return set;
}

function inRange(range: string, chars: Set<number>): boolean {
  const parts = range.split(',').map((p) => {
    const m = /U\+([0-9A-F?]+)(?:-([0-9A-F]+))?/i.exec(p.trim());
    if (!m) return null;
    const a = parseInt(m[1].replace(/\?/g, '0'), 16);
    const b = m[2] ? parseInt(m[2], 16) : parseInt(m[1].replace(/\?/g, 'F'), 16);
    return [a, b] as const;
  });
  for (const c of chars) for (const p of parts) if (p && c >= p[0] && c <= p[1]) return true;
  return false;
}

const dataUrlCache = new Map<string, Promise<string>>();

function toDataUrl(url: string): Promise<string> {
  let p = dataUrlCache.get(url);
  if (!p) {
    p = fetch(url)
      .then((r) => r.blob())
      .then(
        (b) =>
          new Promise<string>((res) => {
            const fr = new FileReader();
            fr.onload = () => res(String(fr.result));
            fr.readAsDataURL(b);
          }),
      );
    dataUrlCache.set(url, p);
  }
  return p;
}

/** @font-face rules (with embedded files) for the fonts a rendered slide set actually uses. */
export async function embeddedFontCss(root: HTMLElement, pres: Presentation): Promise<string> {
  const families = new Set<string>();
  root.querySelectorAll<HTMLElement>('.sl-slide, .sl-slide *').forEach((el) => {
    const ff = el.style.fontFamily;
    if (ff) for (const f of ff.split(',')) families.add(f.trim().replace(/^['"]|['"]$/g, '').toLowerCase());
  });
  const chars = usedChars(pres);
  const faces = BUNDLED_FACES.filter((f) => families.has(f.family.toLowerCase()) && inRange(f.range, chars));
  const rules = await Promise.all(
    faces.map(async (f) => `@font-face{font-family:${JSON.stringify(f.family)};font-style:${f.style};font-weight:${f.weight};src:url(${await toDataUrl(f.url)}) format('woff2');unicode-range:${f.range};}`),
  );
  return rules.join('\n');
}

/* ------------------------------------------------------------ rendering */

interface Offscreen {
  host: HTMLElement;
  slides: HTMLElement[];
  dispose(): void;
}

async function waitForAssets(host: HTMLElement): Promise<void> {
  await document.fonts.ready;
  const imgs = Array.from(host.querySelectorAll('img'));
  await Promise.all(imgs.map((i) => (i.complete ? undefined : new Promise((r) => ((i.onload = r), (i.onerror = r))))));
  // charts are created asynchronously once Chart.js has loaded
  const charts = host.querySelectorAll('.sl-chart');
  if (charts.length) {
    await import('chart.js/auto');
    for (let i = 0; i < 40; i++) {
      if (Array.from(host.querySelectorAll<HTMLCanvasElement>('.sl-chart canvas')).every((c) => c.dataset.ready === '1')) break;
      await new Promise((r) => setTimeout(r, 50));
    }
  }
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
}

async function renderOffscreen(pres: Presentation, indices: number[]): Promise<Offscreen> {
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-200000px;top:0;pointer-events:none;';
  document.body.appendChild(host);
  const root = createRoot(host);
  flushSync(() =>
    root.render(
      <>
        {indices.map((i) => (
          <SlideView key={i} pres={pres} slide={pres.slides[i]} slideNumber={i + 1} mode="print" />
        ))}
      </>,
    ),
  );
  await waitForAssets(host);
  return {
    host,
    slides: Array.from(host.querySelectorAll<HTMLElement>(':scope > .sl-slide')),
    dispose: () => {
      root.unmount();
      host.remove();
    },
  };
}

/** Static HTML of a rendered slide (charts become pictures). */
function snapshot(node: HTMLElement): string {
  const clone = node.cloneNode(true) as HTMLElement;
  const src = Array.from(node.querySelectorAll('canvas'));
  Array.from(clone.querySelectorAll('canvas')).forEach((c, i) => {
    const img = document.createElement('img');
    try {
      img.src = src[i].toDataURL('image/png');
    } catch {
      img.alt = '';
    }
    img.style.cssText = 'width:100%;height:100%;display:block';
    c.replaceWith(img);
  });
  return clone.outerHTML;
}

/* ------------------------------------------------------------------ PNG */

async function nodeToPng(node: HTMLElement, pres: Presentation, fontCss: string, pixelRatio: number): Promise<Uint8Array> {
  const { toPng } = await import('html-to-image');
  const url = await toPng(node, { width: pres.size.w, height: pres.size.h, pixelRatio, fontEmbedCSS: fontCss, cacheBust: false, skipAutoScale: true });
  const b64 = url.slice(url.indexOf(',') + 1);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export async function slidePng(pres: Presentation, index: number, width = 1920): Promise<Uint8Array> {
  const off = await renderOffscreen(pres, [index]);
  try {
    const css = await embeddedFontCss(off.host, pres);
    return await nodeToPng(off.slides[0], pres, css, width / pres.size.w);
  } finally {
    off.dispose();
  }
}

export async function slidesZip(pres: Presentation, title: string, width = 1920): Promise<Uint8Array> {
  const idx = pres.slides.map((_, i) => i);
  const off = await renderOffscreen(pres, idx);
  try {
    const css = await embeddedFontCss(off.host, pres);
    const files: Record<string, Uint8Array> = {};
    const base = title.replace(/\.[^.]+$/, '') || 'Slides';
    for (let i = 0; i < off.slides.length; i++) files[`${base} - Slide ${String(i + 1).padStart(2, '0')}.png`] = await nodeToPng(off.slides[i], pres, css, width / pres.size.w);
    return zipSync(files, { level: 0 });
  } finally {
    off.dispose();
  }
}

/** A thumbnail as a data URL (used for file previews and exported icons). */
export async function slideThumbDataUrl(pres: Presentation, index: number, width = 320): Promise<string> {
  const off = await renderOffscreen(pres, [index]);
  try {
    const { toPng } = await import('html-to-image');
    const css = await embeddedFontCss(off.host, pres);
    return await toPng(off.slides[0], { width: pres.size.w, height: pres.size.h, pixelRatio: width / pres.size.w, fontEmbedCSS: css, skipAutoScale: true });
  } finally {
    off.dispose();
  }
}

/* ------------------------------------------------------------ PDF / HTML */

export async function slidesHtml(pres: Presentation, indices = pres.slides.map((_, i) => i)): Promise<{ slides: string[]; fontCss: string }> {
  const off = await renderOffscreen(pres, indices);
  try {
    const fontCss = await embeddedFontCss(off.host, pres);
    return { slides: off.slides.map(snapshot), fontCss };
  } finally {
    off.dispose();
  }
}

export async function buildSlidesPdfJob(pres: Presentation, title: string): Promise<PdfJob> {
  const indices = pres.slides.map((s, i) => (s.hidden ? -1 : i)).filter((i) => i >= 0);
  const { slides } = await slidesHtml(pres, indices.length ? indices : [0]);
  const w = pres.size.w / 96;
  const h = pres.size.h / 96;
  return {
    title,
    bodyClass: 'print-slides',
    pageWidthIn: w,
    pageHeightIn: h,
    marginsIn: { top: 0, right: 0, bottom: 0, left: 0 },
    css: `${slidesCss}
.print-slides{margin:0}
.pdf-page{width:${pres.size.w}px;height:${pres.size.h}px;overflow:hidden;break-after:page;page-break-after:always;position:relative}
.pdf-page:last-child{break-after:auto;page-break-after:auto}
.pdf-page .sl-slide{-webkit-print-color-adjust:exact;print-color-adjust:exact}`,
    html: slides.map((s) => `<div class="pdf-page">${s}</div>`).join('\n'),
  };
}
