import type { ThemeSwatch } from '@/ui/color';
import { colorRef, parseColorRef, resolveColor, rgbToHsl, type Theme, type ThemeColorName } from '../model';

const NAMES: [ThemeColorName, string][] = [
  ['bg1', 'Background 1'],
  ['tx1', 'Text 1'],
  ['bg2', 'Background 2'],
  ['tx2', 'Text 2'],
  ['accent1', 'Accent 1'],
  ['accent2', 'Accent 2'],
  ['accent3', 'Accent 3'],
  ['accent4', 'Accent 4'],
  ['accent5', 'Accent 5'],
  ['accent6', 'Accent 6'],
];

function lightness(hex: string): number {
  const n = parseInt(hex.replace('#', '').slice(0, 6), 16);
  return rgbToHsl((n >> 16) & 255, (n >> 8) & 255, n & 255)[2];
}

/** Office-style theme colour grid: base colours plus lighter/darker variants, all as theme references. */
export function themeSwatches(theme: Theme): ThemeSwatch[][] {
  const rows: ThemeSwatch[][] = [NAMES.map(([n, label]) => ({ color: resolveColor(`@${n}`, theme), value: `@${n}`, title: label }))];
  for (let r = 1; r <= 5; r++) {
    rows.push(
      NAMES.map(([n, label]) => {
        const l = lightness(theme.colors[n] ?? '#808080');
        const lum = l > 0.9 ? -[5, 15, 25, 35, 50][r - 1] : l < 0.1 ? [50, 35, 25, 15, 5][r - 1] : r <= 3 ? [80, 60, 40][r - 1] : -[25, 50][r - 4];
        const v = colorRef(n, lum);
        return { color: resolveColor(v, theme), value: v, title: `${label}, ${lum > 0 ? 'lighter' : 'darker'} ${Math.abs(lum)}%` };
      }),
    );
  }
  return rows;
}

export function describeColor(v: string): string {
  const ref = parseColorRef(v);
  if (!ref) return v.toUpperCase();
  const label = NAMES.find(([n]) => n === ref.name)?.[1] ?? ref.name;
  const lum = ref.lum ? `, ${ref.lum > 0 ? 'lighter' : 'darker'} ${Math.abs(ref.lum)}%` : '';
  const alpha = ref.alpha < 1 ? `, ${Math.round((1 - ref.alpha) * 100)}% transparent` : '';
  return `${label}${lum}${alpha}`;
}

export function colorProps(theme: Theme) {
  return { themeSwatches: themeSwatches(theme), display: (v: string) => resolveColor(v, theme) };
}
