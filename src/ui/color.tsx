import { ChevronDown, Pipette, Ban } from 'lucide-react';
import { useRef, useState, type ReactNode } from 'react';
import { create } from 'zustand';
import { Popover } from './popover';

/* ------------------------------------------------------------ colour math */

export function hexToRgb(hex: string): [number, number, number] | null {
  let h = hex.trim().replace(/^#/, '');
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  if (h.length === 8) h = h.slice(0, 6);
  if (!/^[0-9a-f]{6}$/i.test(h)) return null;
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

export function rgbToHex(r: number, g: number, b: number): string {
  const c = (v: number) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0');
  return `#${c(r)}${c(g)}${c(b)}`;
}

/** Normalises any CSS colour (named, rgb(), #rgb…) to #rrggbb, or null. */
function parseHexOrRgb(s: string): string | null | undefined {
  const rgb = hexToRgb(s);
  if (rgb) return rgbToHex(...rgb);
  const m = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.]+%?))?\s*\)$/i.exec(s);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
    return rgbToHex(+m[1], +m[2], +m[3]);
  }
  return undefined;
}

let colorCtx: CanvasRenderingContext2D | null | undefined;

/** Normalises any CSS colour (named, rgb(), #rgb…) to #rrggbb, or null for none/transparent. */
export function normalizeColor(input: string | null | undefined): string | null {
  if (!input) return null;
  const s = input.trim();
  if (!s || s === 'transparent' || s === 'none' || s === 'auto' || s === 'inherit') return null;
  const direct = parseHexOrRgb(s);
  if (direct !== undefined) return direct;
  if (typeof document === 'undefined') return null;
  if (colorCtx === undefined) colorCtx = document.createElement('canvas').getContext('2d');
  if (!colorCtx) return null;
  colorCtx.fillStyle = '#010203';
  colorCtx.fillStyle = s;
  const out = String(colorCtx.fillStyle);
  if (out === '#010203') return null; // not a valid colour
  return parseHexOrRgb(out) ?? null;
}

export function mix(hex: string, target: string, amount: number): string {
  const a = hexToRgb(hex) ?? [0, 0, 0];
  const b = hexToRgb(target) ?? [255, 255, 255];
  return rgbToHex(a[0] + (b[0] - a[0]) * amount, a[1] + (b[1] - a[1]) * amount, a[2] + (b[2] - a[2]) * amount);
}

export function luminance(hex: string): number {
  const rgb = hexToRgb(hex);
  if (!rgb) return 1;
  const [r, g, b] = rgb.map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

export const readableOn = (bg: string) => (luminance(bg) > 0.45 ? '#111111' : '#ffffff');

/* --------------------------------------------------------------- palettes */

export const THEME_BASE = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
export const STANDARD_COLORS = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];

export function themeGrid(base = THEME_BASE): string[][] {
  // Row 0: base colours, rows 1-5: tints and shades like Office.
  return [0, 1, 2, 3, 4, 5].map((row) =>
    base.map((c) => {
      if (row === 0) return c;
      const l = luminance(c);
      if (l > 0.9) return mix(c, '#000000', [0.05, 0.15, 0.25, 0.35, 0.5][row - 1]);
      if (l < 0.02) return mix(c, '#ffffff', [0.5, 0.35, 0.25, 0.15, 0.05][row - 1]);
      return row <= 3 ? mix(c, '#ffffff', [0.8, 0.6, 0.4][row - 1]) : mix(c, '#000000', [0.25, 0.5][row - 4]);
    }),
  );
}

const HIGHLIGHTS = ['#ffff00', '#00ff00', '#00ffff', '#ff00ff', '#0000ff', '#ff0000', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#c0c0c0', '#000000'];
export const HIGHLIGHT_COLORS = HIGHLIGHTS;

/* ------------------------------------------------------------ recent store */

const useRecentColors = create<{ colors: string[]; add(c: string): void }>((set) => ({
  colors: (() => {
    try {
      return JSON.parse(localStorage.getItem('affice.recentColors') ?? '[]');
    } catch {
      return [];
    }
  })(),
  add: (c) =>
    set((s) => {
      const colors = [c, ...s.colors.filter((x) => x !== c)].slice(0, 10);
      try {
        localStorage.setItem('affice.recentColors', JSON.stringify(colors));
      } catch {
        /* ignore */
      }
      return { colors };
    }),
}));

/* ------------------------------------------------------------------ picker */

export interface ThemeSwatch {
  /** CSS colour shown on the swatch. */
  color: string;
  /** Value passed to onPick (a colour or a theme reference). */
  value: string;
  title: string;
}

export interface ColorPickerProps {
  value?: string | null;
  onPick: (color: string | null) => void;
  noneLabel?: string;
  palette?: 'theme' | 'highlight';
  themeColors?: string[];
  /** Theme swatch grid that picks references instead of fixed colours (Slides). */
  themeSwatches?: ThemeSwatch[][];
  /** Turns a picked value into a CSS colour (for theme references). */
  display?: (value: string) => string;
}

export function ColorPicker({ value, onPick, noneLabel, palette = 'theme', themeColors, themeSwatches, display = (v) => v }: ColorPickerProps) {
  const recent = useRecentColors();
  const [custom, setCustom] = useState(value && value.startsWith('#') ? value : '#004fff');
  const pick = (c: string | null) => {
    if (c) recent.add(c);
    onPick(c);
  };
  const eyedropper = async () => {
    const ED = (window as unknown as { EyeDropper?: new () => { open(): Promise<{ sRGBHex: string }> } }).EyeDropper;
    if (!ED) return;
    try {
      const r = await new ED().open();
      const c = normalizeColor(r.sRGBHex);
      if (c) pick(c);
    } catch {
      /* cancelled */
    }
  };
  const sw = (c: string, key: string, title?: string, v: string = c) => (
    <button
      key={key}
      type="button"
      className={`swatch${value && value.toLowerCase() === v.toLowerCase() ? ' selected' : ''}`}
      style={{ background: c }}
      title={title ?? c.toUpperCase()}
      onClick={() => pick(v)}
    />
  );
  const grid: ThemeSwatch[][] = themeSwatches ?? themeGrid(themeColors).map((row) => row.map((c) => ({ color: c, value: c, title: c.toUpperCase() })));
  return (
    <div className="color-picker" onMouseDown={(e) => e.target instanceof HTMLInputElement || e.preventDefault()}>
      {noneLabel && (
        <button type="button" className="color-none" onClick={() => pick(null)}>
          <Ban size={15} /> {noneLabel}
        </button>
      )}
      {palette === 'highlight' ? (
        <div className="swatch-grid">{HIGHLIGHTS.map((c, i) => sw(c, `h${i}`))}</div>
      ) : (
        <>
          <div className="color-section">Theme colours</div>
          <div className="swatch-theme">
            {grid.map((row, r) => (
              <div key={r} className={`swatch-row${r === 0 ? ' base' : ''}`}>
                {row.map((c, i) => sw(c.color, `${r}-${i}`, c.title, c.value))}
              </div>
            ))}
          </div>
          <div className="color-section">Standard colours</div>
          <div className="swatch-row">{STANDARD_COLORS.map((c, i) => sw(c, `s${i}`))}</div>
        </>
      )}
      {recent.colors.length > 0 && (
        <>
          <div className="color-section">Recent</div>
          <div className="swatch-row">{recent.colors.map((c, i) => sw(display(c), `r${i}`, undefined, c))}</div>
        </>
      )}
      <div className="color-custom">
        <input type="color" value={custom} onChange={(e) => setCustom(e.target.value)} aria-label="Custom colour" />
        <input
          className="input input-sm"
          value={custom}
          onChange={(e) => setCustom(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              const c = normalizeColor(custom);
              if (c) pick(c);
            }
          }}
          aria-label="Hex colour"
          style={{ width: 84 }}
        />
        <button type="button" className="btn btn-sm" onClick={() => {
          const c = normalizeColor(custom);
          if (c) pick(c);
        }}>
          Apply
        </button>
        {'EyeDropper' in window && (
          <button type="button" className="icon-btn icon-btn-sm" data-tip="Pick a colour from the screen" onClick={eyedropper}>
            <Pipette size={15} />
          </button>
        )}
      </div>
    </div>
  );
}

/** Split button: main part applies the current colour, arrow opens the picker. */
export function ColorSplitButton({
  icon,
  color,
  onApply,
  tip,
  noneLabel,
  palette,
  big,
  label,
  themeSwatches,
  display = (v: string) => v,
  extra,
}: {
  icon: ReactNode;
  color: string | null;
  onApply: (c: string | null) => void;
  tip: string;
  noneLabel?: string;
  palette?: 'theme' | 'highlight';
  big?: boolean;
  label?: string;
  themeSwatches?: ThemeSwatch[][];
  display?: (value: string) => string;
  /** Extra content under the picker (e.g. "More fill options…"). */
  extra?: (close: () => void) => ReactNode;
}) {
  const [open, setOpen] = useState(false);
  const wrap = useRef<HTMLDivElement>(null);
  const [current, setCurrent] = useState<string | null>(color);
  return (
    <div ref={wrap} className={`split-btn color-split${big ? ' big' : ''}${open ? ' open' : ''}`}>
      <button type="button" className="split-main" data-tip={tip} onMouseDown={(e) => e.preventDefault()} onClick={() => onApply(current)}>
        <span className="color-split-icon">
          {icon}
          <span className="color-bar" style={{ background: current ? display(current) : 'transparent', borderColor: current ? display(current) : 'var(--border-strong)' }} />
        </span>
        {label && <span className="split-label">{label}</span>}
      </button>
      <button type="button" className="split-arrow" aria-label={`${tip} options`} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen((o) => !o)}>
        <ChevronDown size={12} />
      </button>
      <Popover open={open} anchor={wrap.current} onClose={() => setOpen(false)} ignore={[wrap.current]}>
        <ColorPicker
          value={current}
          palette={palette}
          noneLabel={noneLabel}
          themeSwatches={themeSwatches}
          display={display}
          onPick={(c) => {
            setCurrent(c);
            onApply(c);
            setOpen(false);
          }}
        />
        {extra?.(() => setOpen(false))}
      </Popover>
    </div>
  );
}

/** Button showing a colour swatch that opens a picker (for panels/dialogs). */
export function ColorField({
  value,
  onChange,
  noneLabel,
  label,
  themeSwatches,
  display = (v: string) => v,
  describe,
}: {
  value: string | null;
  onChange: (c: string | null) => void;
  noneLabel?: string;
  label?: string;
  themeSwatches?: ThemeSwatch[][];
  display?: (value: string) => string;
  /** Text shown for the current value (defaults to the hex code). */
  describe?: (value: string) => string;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLButtonElement>(null);
  return (
    <>
      <button ref={ref} type="button" className="color-field" aria-label={label} onClick={() => setOpen((o) => !o)}>
        <span className={`color-field-swatch${value ? '' : ' none'}`} style={{ background: value ? display(value) : undefined }} />
        <span className="color-field-text">{value ? (describe ? describe(value) : value.toUpperCase()) : (noneLabel ?? 'None')}</span>
        <ChevronDown size={13} />
      </button>
      <Popover open={open} anchor={ref.current} onClose={() => setOpen(false)} ignore={[ref.current]}>
        <ColorPicker
          value={value}
          noneLabel={noneLabel}
          themeSwatches={themeSwatches}
          display={display}
          onPick={(c) => {
            onChange(c);
            setOpen(false);
          }}
        />
      </Popover>
    </>
  );
}
