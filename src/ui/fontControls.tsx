import { useEffect, useState } from 'react';
import { allFontOptions, FONT_SIZES, loadSystemFonts } from '@/lib/fonts';
import { ComboBox, type ComboOption } from './controls';

const GROUP_LABEL = { office: 'Office-compatible', bundled: 'Affice fonts', system: 'Installed on this computer' } as const;

export function FontFamilyCombo({ value, onChange, width = 150 }: { value: string; onChange: (f: string) => void; width?: number }) {
  const [, force] = useState(0);
  useEffect(() => {
    void loadSystemFonts().then(() => force((n) => n + 1));
  }, []);
  const options = (): ComboOption[] =>
    allFontOptions().map((f) => ({
      value: f.name,
      group: GROUP_LABEL[f.source],
      style: { fontFamily: `"${f.name}", sans-serif`, fontSize: 14 },
      hint: f.source === 'office' ? 'metric-compatible' : undefined,
    }));
  return (
    <ComboBox
      value={value}
      options={options}
      onCommit={onChange}
      width={width}
      label="Font"
      listWidth={300}
      className="font-combo"
    />
  );
}

export function nextFontSize(current: number, dir: 1 | -1): number {
  if (dir > 0) return FONT_SIZES.find((s) => s > current + 0.01) ?? Math.min(1638, Math.round(current + 10));
  const smaller = [...FONT_SIZES].reverse().find((s) => s < current - 0.01);
  return smaller ?? Math.max(1, Math.round(current - 1));
}

export function FontSizeCombo({ value, onChange, width = 62 }: { value: number | null; onChange: (pt: number) => void; width?: number }) {
  const shown = value === null ? '' : String(Math.round(value * 10) / 10);
  return (
    <ComboBox
      value={shown}
      options={FONT_SIZES.map((s) => ({ value: String(s) }))}
      filter={false}
      numeric
      onStep={(d) => value !== null && onChange(nextFontSize(value, d))}
      onCommit={(v) => {
        const n = parseFloat(v);
        if (Number.isFinite(n) && n >= 1 && n <= 1638) onChange(n);
      }}
      width={width}
      label="Font size"
    />
  );
}
