import type { JSONContent } from '@tiptap/core';
import { dataUrlToBytes } from '@/lib/utils';
import { renderMathToHtml } from '../editor/extensions';

export function textOf(n: JSONContent): string {
  if (n.type === 'text') return n.text ?? '';
  if (n.type === 'hardBreak') return '\n';
  return (n.content ?? []).map(textOf).join('');
}

/**
 * Headings in document order. With `topLevelOnly`, only top-level headings are returned
 * (including empty ones) so indexes line up with the pagination's heading→page list.
 */
export function collectHeadingsFromJson(doc: JSONContent, maxLevel = 3, topLevelOnly = false): Array<{ level: number; text: string }> {
  const out: Array<{ level: number; text: string }> = [];
  if (topLevelOnly) {
    for (const n of doc.content ?? []) {
      if (n.type !== 'heading') continue;
      const level = Number(n.attrs?.level) || 1;
      if (level <= maxLevel) out.push({ level, text: textOf(n).trim() });
    }
    return out;
  }
  const walk = (n: JSONContent) => {
    if (n.type === 'heading') {
      const level = Number(n.attrs?.level) || 1;
      const text = textOf(n).trim();
      if (level <= maxLevel && text) out.push({ level, text });
      return;
    }
    if (n.type === 'table') return;
    n.content?.forEach(walk);
  };
  walk(doc);
  return out;
}

export function slugifyHeading(text: string, used: Map<string, number>): string {
  const base =
    text
      .toLowerCase()
      .normalize('NFKD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 36) || 'section';
  const n = used.get(base) ?? 0;
  used.set(base, n + 1);
  // Word bookmark names must start with a letter
  return `_${n ? `${base}-${n + 1}` : base}`.replace(/-/g, '_');
}

/** Renders a LaTeX formula to a PNG (2× resolution) for formats without native maths support. */
export async function renderMathPng(latex: string): Promise<{ data: Uint8Array; w: number; h: number } | null> {
  try {
    const { toPng } = await import('html-to-image');
    const host = document.createElement('div');
    host.style.cssText = 'position:fixed;left:-10000px;top:0;padding:2px 3px;background:#fff;color:#000;font-size:14.67px;display:inline-block';
    host.innerHTML = renderMathToHtml(latex, false);
    document.body.appendChild(host);
    await document.fonts.ready;
    const rect = host.getBoundingClientRect();
    const url = await toPng(host, { pixelRatio: 2, backgroundColor: '#ffffff', cacheBust: false });
    host.remove();
    const parsed = dataUrlToBytes(url);
    if (!parsed) return null;
    return { data: parsed.bytes, w: Math.max(1, Math.round(rect.width)), h: Math.max(1, Math.round(rect.height)) };
  } catch {
    return null;
  }
}
