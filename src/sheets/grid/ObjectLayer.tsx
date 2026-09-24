import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { Chart as ChartJS } from 'chart.js';
import { useSettings } from '@/app/settings';
import type { ChartSpec, ImageSpec } from '../model/types';
import type { SheetDoc } from '../doc';
import type { SheetUI } from '../ui/controller';
import { upsertChart, upsertImage } from '../ops/data';
import { buildChartConfig } from '../charts/chartConfig';
import type { Mapper } from './render';

type Obj = { kind: 'chart'; spec: ChartSpec } | { kind: 'image'; spec: ImageSpec };

let chartMod: Promise<typeof import('chart.js/auto')> | null = null;
const loadChart = () => (chartMod ??= import('chart.js/auto'));

function ChartCanvas({ doc, spec, version }: { doc: SheetDoc; spec: ChartSpec; version: number }) {
  const ref = useRef<HTMLCanvasElement>(null);
  const chart = useRef<ChartJS | null>(null);
  const motion = useSettings((s) => s.settings.motion);
  const dark = document.documentElement.getAttribute('data-theme') === 'dark';
  useEffect(() => {
    let dead = false;
    void loadChart().then((mod) => {
      if (dead || !ref.current) return;
      const animate = motion === 'full';
      const cfg = buildChartConfig(doc, spec, doc.sheet, dark, animate);
      const prevType = chart.current ? (chart.current.config as { type?: string }).type : undefined;
      if (chart.current && prevType === (cfg as { type?: string }).type) {
        chart.current.data = cfg.data;
        chart.current.options = cfg.options ?? {};
        chart.current.update(animate ? undefined : 'none');
      } else {
        chart.current?.destroy();
        chart.current = new mod.default(ref.current, cfg);
      }
    });
    return () => {
      dead = true;
    };
  }, [doc, spec, version, dark, motion]);
  useEffect(
    () => () => {
      chart.current?.destroy();
      chart.current = null;
    },
    [],
  );
  return <canvas ref={ref} />;
}

export function ObjectLayer({ doc, ui, mapper, width, height, headerW, headerH }: { doc: SheetDoc; ui: SheetUI; mapper: Mapper; width: number; height: number; headerW: number; headerH: number }) {
  const docVersion = useSyncExternalStore(doc.subscribe, doc.getVersion);
  useSyncExternalStore(ui.subscribe, ui.getVersion);
  const [drag, setDrag] = useState<{ id: string; kind: 'chart' | 'image'; mode: 'move' | string; sx: number; sy: number; x: number; y: number; w: number; h: number; ox: number; oy: number; ow: number; oh: number } | null>(null);
  const sheet = doc.sheet;
  const z = ui.zoom;
  const objs: Obj[] = [...sheet.images.map((spec) => ({ kind: 'image' as const, spec })), ...sheet.charts.map((spec) => ({ kind: 'chart' as const, spec }))];
  if (!objs.length) return null;

  const place = (o: Obj) => {
    const a = o.spec.anchor;
    const x = mapper.x(a.c) + a.dx * z;
    const y = mapper.y(a.r) + a.dy * z;
    return { x, y, w: o.spec.w * z, h: o.spec.h * z };
  };

  const commit = (o: Obj, x: number, y: number, w: number, h: number) => {
    const pos = mapper.cellAt(Math.max(headerW + 1, x), Math.max(headerH + 1, y));
    const anchor = { r: pos.r, c: pos.c, dx: Math.max(0, (x - mapper.x(pos.c)) / z), dy: Math.max(0, (y - mapper.y(pos.r)) / z) };
    const next = { ...o.spec, anchor, w: Math.max(40, Math.round(w / z)), h: Math.max(30, Math.round(h / z)) };
    if (o.kind === 'chart') upsertChart(doc, next as ChartSpec);
    else upsertImage(doc, next as ImageSpec);
  };

  const start = (e: React.PointerEvent, o: Obj, mode: string) => {
    e.stopPropagation();
    e.preventDefault();
    ui.selectedObject = { kind: o.kind, id: o.spec.id };
    ui.emit();
    const p = place(o);
    setDrag({ id: o.spec.id, kind: o.kind, mode, sx: e.clientX, sy: e.clientY, ...p, ox: p.x, oy: p.y, ow: p.w, oh: p.h });
    (e.target as HTMLElement).setPointerCapture(e.pointerId);
  };

  const move = (e: React.PointerEvent) => {
    if (!drag) return;
    const dx = e.clientX - drag.sx;
    const dy = e.clientY - drag.sy;
    let { ox: x, oy: y, ow: w, oh: h } = drag;
    if (drag.mode === 'move') {
      x += dx;
      y += dy;
    } else {
      if (drag.mode.includes('e')) w = Math.max(40, w + dx);
      if (drag.mode.includes('s')) h = Math.max(30, h + dy);
      if (drag.mode.includes('w')) {
        w = Math.max(40, w - dx);
        x = drag.ox + drag.ow - w;
      }
      if (drag.mode.includes('n')) {
        h = Math.max(30, h - dy);
        y = drag.oy + drag.oh - h;
      }
      if (e.shiftKey && drag.kind === 'image') {
        const ratio = drag.ow / drag.oh;
        h = w / ratio;
      }
    }
    setDrag({ ...drag, x, y, w, h });
  };

  const end = (o: Obj) => {
    if (!drag) return;
    if (drag.x !== drag.ox || drag.y !== drag.oy || drag.w !== drag.ow || drag.h !== drag.oh) commit(o, drag.x, drag.y, drag.w, drag.h);
    setDrag(null);
  };

  return (
    <div className="object-layer" style={{ left: headerW, top: headerH, width: Math.max(0, width - headerW - 14), height: Math.max(0, height - headerH - 14) }}>
      {objs.map((o) => {
        const p = drag && drag.id === o.spec.id ? { x: drag.x, y: drag.y, w: drag.w, h: drag.h } : place(o);
        const selected = ui.selectedObject?.id === o.spec.id;
        if (p.x + p.w < headerW || p.y + p.h < headerH || p.x > width || p.y > height) return null;
        return (
          <div
            key={o.spec.id}
            className={`sheet-object ${o.kind}${selected ? ' selected' : ''}`}
            style={{ left: p.x - headerW, top: p.y - headerH, width: p.w, height: p.h }}
            onPointerDown={(e) => start(e, o, 'move')}
            onPointerMove={move}
            onPointerUp={() => end(o)}
            onDoubleClick={() => o.kind === 'chart' && ui.onEditChart?.(o.spec.id)}
            onContextMenu={(e) => {
              e.stopPropagation();
            }}
          >
            {o.kind === 'chart' ? <ChartCanvas doc={doc} spec={o.spec} version={docVersion} /> : <img src={o.spec.src} alt={o.spec.alt ?? ''} draggable={false} />}
            {selected &&
              ['n', 's', 'e', 'w', 'ne', 'nw', 'se', 'sw'].map((hd) => (
                <span key={hd} className={`obj-handle h-${hd}`} onPointerDown={(e) => start(e, o, hd)} onPointerMove={move} onPointerUp={() => end(o)} />
              ))}
          </div>
        );
      })}
    </div>
  );
}
