import { BUNDLED_FACES, BUNDLED_FAMILIES } from './fonts.generated';

let injected = false;

/** Registers @font-face rules for all bundled fonts and Office-font aliases. Fonts download lazily when first used. */
export function installFonts(): void {
  if (injected) return;
  injected = true;
  const css = BUNDLED_FACES.map((f) => {
    const src = [...f.locals.map((l) => `local(${JSON.stringify(l)})`), `url(${JSON.stringify(f.url)}) format('woff2')`].join(', ');
    return `@font-face{font-family:${JSON.stringify(f.family)};font-style:${f.style};font-weight:${f.weight};font-display:swap;src:${src};unicode-range:${f.range};}`;
  }).join('\n');
  const style = document.createElement('style');
  style.id = 'affice-fonts';
  style.textContent = css;
  document.head.appendChild(style);
}

export interface FontOption {
  name: string;
  category: 'sans' | 'serif' | 'mono' | 'display' | 'script' | 'system';
  source: 'office' | 'bundled' | 'system';
}

/** Office-compatible names first (resolve to the real font or a metric-compatible clone). */
const OFFICE_FONTS: FontOption[] = [
  { name: 'Calibri', category: 'sans', source: 'office' },
  { name: 'Cambria', category: 'serif', source: 'office' },
  { name: 'Arial', category: 'sans', source: 'office' },
  { name: 'Times New Roman', category: 'serif', source: 'office' },
  { name: 'Courier New', category: 'mono', source: 'office' },
  { name: 'Georgia', category: 'serif', source: 'office' },
];

const ALIAS_TARGETS = new Set(BUNDLED_FAMILIES.filter((f) => f.aliases.length).map((f) => f.name));

export const BUNDLED_FONT_OPTIONS: FontOption[] = BUNDLED_FAMILIES.filter((f) => !ALIAS_TARGETS.has(f.name)).map((f) => ({
  name: f.name,
  category: f.category as FontOption['category'],
  source: 'bundled' as const,
}));

export const CORE_FONTS: FontOption[] = [...OFFICE_FONTS, ...BUNDLED_FONT_OPTIONS];

let systemFonts: string[] | null = null;
let systemFontsPromise: Promise<string[]> | null = null;

/** Lists installed system font families (Local Font Access API; granted automatically in the desktop app). */
export function loadSystemFonts(): Promise<string[]> {
  if (systemFonts) return Promise.resolve(systemFonts);
  if (systemFontsPromise) return systemFontsPromise;
  systemFontsPromise = (async () => {
    try {
      const q = (window as unknown as { queryLocalFonts?: () => Promise<Array<{ family: string }>> }).queryLocalFonts;
      if (!q) return [];
      const fonts = await q();
      const set = new Set<string>();
      for (const f of fonts) set.add(f.family);
      systemFonts = Array.from(set).sort((a, b) => a.localeCompare(b));
      return systemFonts;
    } catch {
      systemFonts = [];
      return [];
    }
  })();
  return systemFontsPromise;
}

export function allFontOptions(): FontOption[] {
  const known = new Set(CORE_FONTS.map((f) => f.name.toLowerCase()));
  const sys = (systemFonts ?? []).filter((n) => !known.has(n.toLowerCase())).map((name) => ({ name, category: 'system' as const, source: 'system' as const }));
  return [...CORE_FONTS, ...sys];
}

/** CSS font-family value with sensible generic fallbacks. */
export function cssFontStack(name: string | undefined | null): string {
  if (!name) return 'inherit';
  const n = name.replace(/^['"]|['"]$/g, '');
  const opt = CORE_FONTS.find((f) => f.name.toLowerCase() === n.toLowerCase());
  const generic = opt?.category === 'serif' ? 'serif' : opt?.category === 'mono' ? 'monospace' : 'sans-serif';
  return `${JSON.stringify(n)}, ${generic}`;
}

/** Ensures the given font faces are loaded (useful before measuring or drawing text on a canvas). */
export async function ensureFontsLoaded(specs: string[]): Promise<void> {
  try {
    await Promise.all(specs.map((s) => document.fonts.load(s)));
  } catch {
    /* ignore */
  }
}

export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 96];
