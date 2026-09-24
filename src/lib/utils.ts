export function uid(prefix = ''): string {
  const rnd = Math.random().toString(36).slice(2, 10);
  return `${prefix}${Date.now().toString(36)}${rnd}`;
}

export const clamp = (v: number, min: number, max: number) => (v < min ? min : v > max ? max : v);

export function debounce<A extends unknown[]>(fn: (...args: A) => void, ms: number) {
  let t: ReturnType<typeof setTimeout> | undefined;
  const d = (...args: A) => {
    if (t) clearTimeout(t);
    t = setTimeout(() => {
      t = undefined;
      fn(...args);
    }, ms);
  };
  d.cancel = () => t && clearTimeout(t);
  d.flush = (...args: A) => {
    if (t) clearTimeout(t);
    t = undefined;
    fn(...args);
  };
  return d;
}

export function throttle<A extends unknown[]>(fn: (...args: A) => void, ms: number) {
  let last = 0;
  let t: ReturnType<typeof setTimeout> | undefined;
  let lastArgs: A;
  return (...args: A) => {
    lastArgs = args;
    const now = Date.now();
    const remaining = ms - (now - last);
    if (remaining <= 0) {
      last = now;
      fn(...args);
    } else if (!t) {
      t = setTimeout(() => {
        last = Date.now();
        t = undefined;
        fn(...lastArgs);
      }, remaining);
    }
  };
}

/** Coalesces calls to once per animation frame. */
export function rafThrottle<A extends unknown[]>(fn: (...args: A) => void) {
  let frame = 0;
  let lastArgs: A;
  const r = (...args: A) => {
    lastArgs = args;
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      fn(...lastArgs);
    });
  };
  r.cancel = () => {
    if (frame) cancelAnimationFrame(frame);
    frame = 0;
  };
  return r;
}

export function basename(p: string): string {
  return p.split(/[\\/]/).pop() ?? p;
}

export function dirname(p: string): string {
  const parts = p.split(/[\\/]/);
  parts.pop();
  return parts.join(p.includes('\\') ? '\\' : '/');
}

export function extname(p: string): string {
  const b = basename(p);
  const i = b.lastIndexOf('.');
  return i > 0 ? b.slice(i + 1).toLowerCase() : '';
}

export function stripExt(name: string): string {
  const i = name.lastIndexOf('.');
  return i > 0 ? name.slice(0, i) : name;
}

export function formatBytes(n: number): string {
  if (!Number.isFinite(n)) return '';
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  if (n < 1024 * 1024 * 1024) return `${(n / 1024 / 1024).toFixed(1)} MB`;
  return `${(n / 1024 / 1024 / 1024).toFixed(2)} GB`;
}

export function timeAgo(ts: number): string {
  const s = Math.round((Date.now() - ts) / 1000);
  if (s < 45) return 'Just now';
  const m = Math.round(s / 60);
  if (m < 60) return `${m} min ago`;
  const h = Math.round(m / 60);
  if (h < 24) return `${h} hr ago`;
  const d = Math.round(h / 24);
  if (d === 1) return 'Yesterday';
  if (d < 7) return `${d} days ago`;
  return new Date(ts).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

export function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
}

export function escapeXml(s: string): string {
  return s
    .replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]!)
    // strip characters that are illegal in XML 1.0
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
}

export const encoder = new TextEncoder();
export const decoder = new TextDecoder();

export function bytesToBase64(bytes: Uint8Array): string {
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)));
  }
  return btoa(bin);
}

export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function dataUrlToBytes(url: string): { mime: string; bytes: Uint8Array } | null {
  const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(url);
  if (!m) return null;
  const mime = m[1] || 'application/octet-stream';
  const bytes = m[2] ? base64ToBytes(m[3]) : encoder.encode(decodeURIComponent(m[3]));
  return { mime, bytes };
}

export function bytesToDataUrl(bytes: Uint8Array, mime: string): string {
  return `data:${mime};base64,${bytesToBase64(bytes)}`;
}

export function mimeFromExt(ext: string): string {
  const e = ext.toLowerCase().replace('.', '');
  const map: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    jpe: 'image/jpeg',
    gif: 'image/gif',
    webp: 'image/webp',
    bmp: 'image/bmp',
    svg: 'image/svg+xml',
    avif: 'image/avif',
    ico: 'image/x-icon',
    tif: 'image/tiff',
    tiff: 'image/tiff',
    emf: 'image/emf',
    wmf: 'image/wmf',
  };
  return map[e] ?? 'application/octet-stream';
}

export function extFromMime(mime: string): string {
  const map: Record<string, string> = {
    'image/png': 'png',
    'image/jpeg': 'jpeg',
    'image/gif': 'gif',
    'image/webp': 'webp',
    'image/bmp': 'bmp',
    'image/svg+xml': 'svg',
    'image/avif': 'avif',
  };
  return map[mime] ?? 'png';
}

/** Reads image dimensions from a data URL (browser only). */
export function imageSize(src: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 300, height: img.naturalHeight || 200 });
    img.onerror = () => resolve({ width: 300, height: 200 });
    img.src = src;
  });
}

/** Converts any browser-decodable image to PNG bytes (used for formats Office can't read, e.g. webp/avif). */
export async function rasterizeToPng(src: string, maxSide = 4096): Promise<{ dataUrl: string; width: number; height: number }> {
  const img = new Image();
  img.src = src;
  await img.decode();
  let w = img.naturalWidth || 300;
  let h = img.naturalHeight || 200;
  const scale = Math.min(1, maxSide / Math.max(w, h));
  w = Math.round(w * scale);
  h = Math.round(h * scale);
  const c = document.createElement('canvas');
  c.width = w;
  c.height = h;
  c.getContext('2d')!.drawImage(img, 0, 0, w, h);
  return { dataUrl: c.toDataURL('image/png'), width: w, height: h };
}

export function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

export function isMod(e: { ctrlKey: boolean; metaKey: boolean }): boolean {
  return e.ctrlKey || e.metaKey;
}

export function shallowEqual(a: Record<string, unknown> | null | undefined, b: Record<string, unknown> | null | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  const ka = Object.keys(a);
  const kb = Object.keys(b);
  if (ka.length !== kb.length) return false;
  for (const k of ka) if (a[k] !== b[k]) return false;
  return true;
}

export function deepClone<T>(v: T): T {
  return typeof structuredClone === 'function' ? structuredClone(v) : (JSON.parse(JSON.stringify(v)) as T);
}

export function pluralize(n: number, word: string, plural = `${word}s`): string {
  return `${n.toLocaleString()} ${n === 1 ? word : plural}`;
}

/** Unit helpers (96 CSS px per inch). */
export const PX_PER_IN = 96;
export const PT_PER_IN = 72;
export const EMU_PER_IN = 914400;
export const EMU_PER_PX = EMU_PER_IN / PX_PER_IN; // 9525
export const TWIP_PER_IN = 1440;
export const ptToPx = (pt: number) => (pt * PX_PER_IN) / PT_PER_IN;
export const pxToPt = (px: number) => (px * PT_PER_IN) / PX_PER_IN;
export const inToPx = (inch: number) => inch * PX_PER_IN;
export const twipToPx = (tw: number) => (tw / TWIP_PER_IN) * PX_PER_IN;
export const pxToTwip = (px: number) => Math.round((px / PX_PER_IN) * TWIP_PER_IN);
export const emuToPx = (emu: number) => emu / EMU_PER_PX;
export const pxToEmu = (px: number) => Math.round(px * EMU_PER_PX);
export const cmToIn = (cm: number) => cm / 2.54;
export const inToCm = (i: number) => i * 2.54;
