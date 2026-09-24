import { Minus, Plus } from 'lucide-react';
import { useRef, useState } from 'react';
import { Slider } from '@/ui/controls';
import { Menu } from '@/ui/menu';

export function ZoomControl({
  zoom,
  onZoom,
  min = 0.25,
  max = 4,
  presets = [0.5, 0.75, 1, 1.25, 1.5, 2],
  extra,
}: {
  zoom: number;
  onZoom: (z: number) => void;
  min?: number;
  max?: number;
  presets?: number[];
  extra?: Array<{ label: string; onSelect: () => void }>;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLButtonElement>(null);
  const clampZ = (z: number) => Math.min(max, Math.max(min, Math.round(z * 100) / 100));
  // slider maps 0..100 to min..max around 100% in the middle
  const toSlider = (z: number) => (z <= 1 ? ((z - min) / (1 - min)) * 50 : 50 + ((z - 1) / (max - 1)) * 50);
  const fromSlider = (v: number) => (v <= 50 ? min + (v / 50) * (1 - min) : 1 + ((v - 50) / 50) * (max - 1));
  return (
    <div className="zoom-control">
      <button type="button" className="status-item" data-tip="Zoom out" onClick={() => onZoom(clampZ(zoom - 0.1))}>
        <Minus size={13} />
      </button>
      <Slider value={toSlider(zoom)} min={0} max={100} step={1} onChange={(v) => onZoom(clampZ(Math.abs(v - 50) < 2 ? 1 : fromSlider(v)))} label="Zoom" />
      <button type="button" className="status-item" data-tip="Zoom in" onClick={() => onZoom(clampZ(zoom + 0.1))}>
        <Plus size={13} />
      </button>
      <button ref={ref} type="button" className="status-item zoom-value" data-tip="Zoom options" onClick={() => setOpen((o) => !o)}>
        {Math.round(zoom * 100)}%
      </button>
      <Menu
        open={open}
        anchor={ref.current}
        ignore={[ref.current]}
        placement="top-end"
        onClose={() => setOpen(false)}
        items={[
          ...(extra ?? []).map((e) => ({ label: e.label, onSelect: e.onSelect })),
          ...(extra?.length ? [{ separator: true }] : []),
          ...presets.map((p) => ({ label: `${Math.round(p * 100)}%`, checked: Math.abs(p - zoom) < 0.005, onSelect: () => onZoom(p) })),
        ]}
      />
    </div>
  );
}

export function SaveState({ dirty, path }: { dirty: boolean; path?: string }) {
  return (
    <span className={`status-item save-state${dirty ? ' dirty' : ''}`} data-tip={path ?? 'Not saved to a file yet'}>
      <span className="dot" />
      {dirty ? 'Unsaved changes' : path ? 'Saved' : 'Not saved'}
    </span>
  );
}
