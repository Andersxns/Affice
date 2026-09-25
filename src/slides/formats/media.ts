/** Picture helpers shared by the presentation writers: decoding sources, file types and raster fallbacks. */
import { strToU8 } from 'fflate';
import { base64ToBytes } from '@/lib/utils';

export interface DecodedImage {
  data: Uint8Array;
  /** File extension: png, jpeg, gif, bmp, svg or tiff. */
  ext: string;
}

export function imageExt(mime: string): string {
  const m = mime.toLowerCase();
  if (m.includes('png')) return 'png';
  if (m.includes('jpeg') || m.includes('jpg')) return 'jpeg';
  if (m.includes('gif')) return 'gif';
  if (m.includes('svg')) return 'svg';
  if (m.includes('bmp')) return 'bmp';
  if (m.includes('tif')) return 'tiff';
  if (m.includes('webp')) return 'webp';
  return 'png';
}

export const IMAGE_MIME: Record<string, string> = { png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', tiff: 'image/tiff' };

/** Bytes and type of an image source (data URL or fetchable URL). Formats office apps can't read become PNG. */
export async function decodeImage(src: string): Promise<DecodedImage | null> {
  const m = /^data:([^;,]+)(;charset=[^;,]+)?(;base64)?,(.*)$/s.exec(src);
  if (!m) {
    try {
      const r = await fetch(src);
      const b = new Uint8Array(await r.arrayBuffer());
      return { data: b, ext: imageExt(r.headers.get('content-type') ?? 'image/png') };
    } catch {
      return null;
    }
  }
  const mime = m[1];
  const data = m[3] ? base64ToBytes(m[4]) : strToU8(decodeURIComponent(m[4]));
  const ext = imageExt(mime);
  if (['png', 'jpeg', 'gif', 'bmp', 'svg', 'tiff'].includes(ext)) return { data, ext };
  // formats office apps can't read (webp, avif…) become PNG
  const png = await rasterizeImage(src);
  return png ? { data: png, ext: 'png' } : null;
}

/** Draws an image into a PNG (in the browser; null elsewhere or on failure). */
export async function rasterizeImage(src: string, w?: number, h?: number): Promise<Uint8Array | null> {
  try {
    if (typeof Image === 'undefined' || typeof document === 'undefined') return null;
    const img = new Image();
    img.src = src;
    await img.decode();
    const cw = w ?? img.naturalWidth ?? 512;
    const ch = h ?? img.naturalHeight ?? 512;
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(cw));
    c.height = Math.max(1, Math.round(ch));
    c.getContext('2d')!.drawImage(img, 0, 0, c.width, c.height);
    const blob = await new Promise<Blob | null>((r) => c.toBlob(r, 'image/png'));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
  } catch {
    return null;
  }
}

/** Pixel size and resolution of a PNG, JPEG or GIF (dpi defaults to 96). */
export function imageSize(b: Uint8Array): { w: number; h: number; dpiX: number; dpiY: number } | null {
  const u32 = (i: number) => ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]) >>> 0;
  const u16 = (i: number) => (b[i] << 8) | b[i + 1];
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    const out = { w: u32(16), h: u32(20), dpiX: 96, dpiY: 96 };
    // pHYs chunk: pixels per metre
    let i = 8;
    while (i + 12 <= b.length) {
      const len = u32(i);
      const type = String.fromCharCode(b[i + 4], b[i + 5], b[i + 6], b[i + 7]);
      if (type === 'pHYs' && b[i + 16] === 1) {
        out.dpiX = Math.round(u32(i + 8) * 0.0254) || 96;
        out.dpiY = Math.round(u32(i + 12) * 0.0254) || 96;
        break;
      }
      if (type === 'IDAT' || type === 'IEND') break;
      i += 12 + len;
    }
    return out;
  }
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    let dpiX = 96;
    let dpiY = 96;
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) {
        i++;
        continue;
      }
      const marker = b[i + 1];
      const len = u16(i + 2);
      if (marker === 0xe0 && b[i + 4] === 0x4a && b[i + 5] === 0x46 && b[i + 11] === 1) {
        dpiX = u16(i + 12) || 96;
        dpiY = u16(i + 14) || 96;
      }
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) return { w: u16(i + 7), h: u16(i + 5), dpiX, dpiY };
      i += 2 + len;
    }
    return null;
  }
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { w: b[6] | (b[7] << 8), h: b[8] | (b[9] << 8), dpiX: 96, dpiY: 96 };
  return null;
}
