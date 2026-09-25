import { useCallback, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore, type CSSProperties } from 'react';
import { isMod } from '@/lib/utils';
import type { SlidesDoc } from '../doc';
import { adjHandles, DEFAULT_ADJ, isConnector, isLineGeom } from '../geometry';
import { cloneEls, findEl, rotatedBounds, unionBounds, type El, type ShapeEl, type TableEl } from '../model';
import { copyEls, clipboardText, deleteEls, duplicateEls, group, insertEls, moveEl, nudge, pasteEls, setFrame, shapeEl, textBox, ungroup } from '../ops';
import { SlideView, type PlaceholderInsertKind } from '../render/SlideView';
import { createTable, insertRow } from '../tables';
import { TextEditSession, bodyOf } from './TextEditSession';
import type { SlidesUI } from './ui';

type Frame = { x: number; y: number; w: number; h: number };
type HandleDir = [number, number];

const HANDLES: { id: string; d: HandleDir }[] = [
  { id: 'nw', d: [-1, -1] },
  { id: 'n', d: [0, -1] },
  { id: 'ne', d: [1, -1] },
  { id: 'e', d: [1, 0] },
  { id: 'se', d: [1, 1] },
  { id: 's', d: [0, 1] },
  { id: 'sw', d: [-1, 1] },
  { id: 'w', d: [-1, 0] },
];

type Drag =
  | { kind: 'move'; sx: number; sy: number; ids: string[]; orig: El[]; moved: boolean; copy: boolean }
  | { kind: 'resize'; d: HandleDir; ids: string[]; orig: El[]; box: Frame; rot: number; aspect: boolean }
  | { kind: 'rotate'; ids: string[]; orig: El[]; cx: number; cy: number; a0: number }
  | { kind: 'adjust'; id: string; orig: ShapeEl; idx: number }
  | { kind: 'endpoint'; id: string; orig: ShapeEl; end: 'start' | 'end' }
  | { kind: 'marquee'; sx: number; sy: number; add: boolean; base: string[] }
  | { kind: 'insert'; sx: number; sy: number; x: number; y: number };

export interface CanvasHandlers {
  onContextMenu(e: React.MouseEvent, target: El | null): void;
  onEditChart(id: string): void;
  onDropFiles(files: File[], at: { x: number; y: number }): void;
  onPasteImage(blob: Blob): void;
  onShortcut(e: React.KeyboardEvent): boolean;
  onStartShow(fromCurrent: boolean): void;
  onPainterApply(ids: string[]): void;
  /** A quick-insert button inside an empty placeholder (the placeholder is selected first). */
  onPlaceholderInsert(kind: PlaceholderInsertKind): void;
}

const cursorFor = (d: HandleDir, rot: number) => {
  const base = Math.atan2(d[1], d[0]) * (180 / Math.PI) + rot;
  const a = ((Math.round(base / 45) * 45) % 180 + 180) % 180;
  return a === 0 ? 'ew-resize' : a === 45 ? 'nwse-resize' : a === 90 ? 'ns-resize' : 'nesw-resize';
};

function lineEnds(e: ShapeEl): { x1: number; y1: number; x2: number; y2: number } {
  const x1 = e.flipH ? e.x + e.w : e.x;
  const x2 = e.flipH ? e.x : e.x + e.w;
  const y1 = e.flipV ? e.y + e.h : e.y;
  const y2 = e.flipV ? e.y : e.y + e.h;
  return { x1, y1, x2, y2 };
}

function fromEnds(e: ShapeEl, x1: number, y1: number, x2: number, y2: number): ShapeEl {
  return { ...e, x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH: x2 < x1 || undefined, flipV: y2 < y1 || undefined };
}

const isLineEl = (e: El) => e.type === 'shape' && isLineGeom(e.geom) && isConnector(e.geom);

export function SlideCanvas({ doc, ui, active, h }: { doc: SlidesDoc; ui: SlidesUI; active: boolean; h: CanvasHandlers }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  useSyncExternalStore(ui.subscribe, ui.getVersion);
  const wrapRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const sinkRef = useRef<HTMLTextAreaElement>(null);
  const dragRef = useRef<Drag | null>(null);
  const [box, setBox] = useState({ w: 800, h: 500 });
  const wheelAcc = useRef(0);
  const pres = doc.pres;
  const slide = doc.slide;

  useLayoutEffect(() => {
    const el = wrapRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setBox({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setBox({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  const pad = 36;
  const fitZoom = Math.max(0.05, Math.min((box.w - pad * 2) / pres.size.w, (box.h - pad * 2) / pres.size.h));
  const zoom = ui.fit ? fitZoom : ui.zoom;
  ui.effectiveZoom = zoom;

  const focusSink = useCallback(() => {
    requestAnimationFrame(() => {
      if (!doc.editing) sinkRef.current?.focus({ preventScroll: true });
    });
  }, [doc]);

  useEffect(() => {
    if (active) focusSink();
  }, [active, focusSink]);

  const toSlide = (clientX: number, clientY: number) => {
    const r = stageRef.current!.getBoundingClientRect();
    return { x: (clientX - r.left) / zoom, y: (clientY - r.top) / zoom };
  };

  /* --------------------------------------------------------- editing */

  const startEdit = useCallback(
    (id: string, cell?: [number, number], caret: SlidesUI['pendingCaret'] = 'end') => {
      const el = findEl(doc.slide.elements, id);
      if (!el || el.locked) return;
      if (el.type === 'shape' && !el.text) doc.updateEls('Add text', [id], (e) => ({ ...(e as ShapeEl), text: { paras: [{ runs: [], align: 'center' }], anchor: 'm' } }));
      if (el.type !== 'shape' && !(el.type === 'table' && cell)) return;
      ui.pendingCaret = caret;
      doc.editing = { el: id, cell };
      doc.setSel({ els: [id] });
    },
    [doc, ui],
  );

  /** Tab / Shift+Tab in a table cell moves to the next / previous cell; Tab in the last cell adds a row. */
  const tabCell = useCallback(
    (dir: 1 | -1) => {
      const ed = doc.editing;
      const t = ed?.cell ? findEl(doc.slide.elements, ed.el) : undefined;
      if (!ed?.cell || !t || t.type !== 'table') return;
      const cells: [number, number][] = [];
      t.rows.forEach((row, r) => row.cells.forEach((c, ci) => !c.merged && cells.push([r, ci])));
      const i = cells.findIndex(([r, c]) => r === ed.cell![0] && c === ed.cell![1]);
      let next = cells[i + dir];
      if (!next && dir > 0) {
        doc.updateEls('Insert row', [t.id], (e) => insertRow(e as TableEl, (e as TableEl).rows.length));
        next = [t.rows.length, 0];
      }
      if (!next) return;
      ui.pendingCaret = 'all';
      doc.editing = { el: t.id, cell: next };
      doc.emit();
    },
    [doc, ui],
  );

  const exitEdit = useCallback(() => {
    const ed = doc.editing;
    if (!ed) return;
    doc.editing = null;
    const el = findEl(doc.slide.elements, ed.el);
    // empty text boxes disappear, like in PowerPoint
    if (el && el.type === 'shape' && el.textbox && !el.ph && !bodyOf(el)?.paras.some((p) => p.runs.some((r) => r.text))) deleteEls(doc, [el.id]);
    else doc.setSel({ els: [ed.el] });
    focusSink();
  }, [doc, focusSink]);

  /* ---------------------------------------------------------- snapping */

  const snapTargets = (exclude: Set<string>) => {
    const xs = [0, pres.size.w / 2, pres.size.w];
    const ys = [0, pres.size.h / 2, pres.size.h];
    for (const e of slide.elements) {
      if (exclude.has(e.id) || e.hidden) continue;
      const b = rotatedBounds(e);
      xs.push(b.x, b.x + b.w / 2, b.x + b.w);
      ys.push(b.y, b.y + b.h / 2, b.y + b.h);
    }
    return { xs, ys };
  };

  const snap1 = (vals: number[], cands: number[], thr: number): { d: number; at: number[] } => {
    let best = Infinity;
    let d = 0;
    for (const v of vals)
      for (const c of cands) {
        const dist = c - v;
        if (Math.abs(dist) < Math.abs(best) && Math.abs(dist) <= thr) {
          best = dist;
          d = dist;
        }
      }
    if (best === Infinity) return { d: 0, at: [] };
    const at = [...new Set(cands.filter((c) => vals.some((v) => Math.abs(v + d - c) < 0.5)))];
    return { d, at };
  };

  /* ---------------------------------------------------------- pointer */

  const topLevelId = (target: HTMLElement): string | null => {
    if (target.closest('.sl-decor')) return null;
    let node = target.closest<HTMLElement>('[data-el]');
    if (!node || !stageRef.current?.contains(node)) return null;
    let parent = node.parentElement?.closest<HTMLElement>('[data-el]');
    while (parent && stageRef.current.contains(parent)) {
      node = parent;
      parent = node.parentElement?.closest<HTMLElement>('[data-el]');
    }
    return node.getAttribute('data-el');
  };

  const onPointerDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    const target = e.target as HTMLElement;
    if (target.closest('.sl-pm')) return; // clicks inside the text editor
    if (doc.editing) exitEdit();
    const p = toSlide(e.clientX, e.clientY);
    const handle = target.closest<HTMLElement>('[data-handle]');
    wrapRef.current?.setPointerCapture(e.pointerId);

    // drawing a new object
    if (ui.insert) {
      dragRef.current = { kind: 'insert', sx: p.x, sy: p.y, x: p.x, y: p.y };
      e.preventDefault();
      return;
    }

    const phButton = target.closest<HTMLElement>('[data-ph-insert]');
    if (phButton) {
      const id = topLevelId(target);
      e.preventDefault();
      if (id) {
        doc.selectEls([id]);
        h.onPlaceholderInsert(phButton.dataset.phInsert as PlaceholderInsertKind);
      }
      return;
    }

    if (handle) {
      const kind = handle.dataset.handle!;
      const sel = doc.selected;
      e.preventDefault();
      if (kind === 'rotate') {
        const b = sel.length === 1 ? sel[0] : unionBounds(sel.map(rotatedBounds))!;
        const cx = b.x + b.w / 2;
        const cy = b.y + b.h / 2;
        dragRef.current = { kind: 'rotate', ids: sel.map((s) => s.id), orig: sel, cx, cy, a0: Math.atan2(p.y - cy, p.x - cx) };
      } else if (kind.startsWith('adj')) {
        const el = sel[0] as ShapeEl;
        dragRef.current = { kind: 'adjust', id: el.id, orig: el, idx: Number(kind.slice(3)) };
      } else if (kind === 'start' || kind === 'end') {
        dragRef.current = { kind: 'endpoint', id: sel[0].id, orig: sel[0] as ShapeEl, end: kind };
      } else {
        const d = HANDLES.find((x) => x.id === kind)!.d;
        const single = sel.length === 1;
        const b = single ? { x: sel[0].x, y: sel[0].y, w: sel[0].w, h: sel[0].h } : unionBounds(sel.map((s) => ({ x: s.x, y: s.y, w: s.w, h: s.h })))!;
        const aspectDefault = single && (sel[0].type === 'image' || sel[0].type === 'chart') && d[0] !== 0 && d[1] !== 0;
        dragRef.current = { kind: 'resize', d, ids: sel.map((s) => s.id), orig: sel, box: b, rot: single ? (sel[0].rot ?? 0) : 0, aspect: aspectDefault };
      }
      return;
    }

    const id = topLevelId(target);
    if (id) {
      const el = findEl(slide.elements, id)!;
      if (ui.painter) {
        h.onPainterApply([id]);
        return;
      }
      let ids = doc.sel.els;
      if (e.shiftKey || isMod(e)) {
        ids = ids.includes(id) ? ids.filter((x) => x !== id) : [...ids, id];
        doc.selectEls(ids);
      } else if (!ids.includes(id)) {
        ids = [id];
        doc.selectEls(ids);
      }
      if (!ids.includes(id)) return;
      const els = ids.map((x) => findEl(slide.elements, x)).filter((x): x is El => !!x && !x.locked);
      if (el.locked) return;
      dragRef.current = { kind: 'move', sx: p.x, sy: p.y, ids: els.map((x) => x.id), orig: els, moved: false, copy: isMod(e) && !e.shiftKey };
      focusSink();
      return;
    }
    // empty area: marquee selection
    const add = e.shiftKey || isMod(e);
    if (!add) doc.selectEls([]);
    dragRef.current = { kind: 'marquee', sx: p.x, sy: p.y, add, base: add ? doc.sel.els : [] };
    focusSink();
  };

  const onPointerMove = (e: React.PointerEvent) => {
    const d = dragRef.current;
    if (!d) return;
    const p = toSlide(e.clientX, e.clientY);
    const thr = ui.snap && !e.altKey ? 6 / zoom : 0;
    switch (d.kind) {
      case 'move': {
        let dx = p.x - d.sx;
        let dy = p.y - d.sy;
        if (!d.moved && Math.hypot(dx, dy) * zoom < 3) return;
        d.moved = true;
        if (e.shiftKey) {
          if (Math.abs(dx) > Math.abs(dy)) dy = 0;
          else dx = 0;
        }
        const b = unionBounds(d.orig.map(rotatedBounds))!;
        const t = snapTargets(new Set(d.ids));
        const sx = thr ? snap1([b.x + dx, b.x + b.w / 2 + dx, b.x + b.w + dx], t.xs, thr) : { d: 0, at: [] };
        const sy = thr ? snap1([b.y + dy, b.y + b.h / 2 + dy, b.y + b.h + dy], t.ys, thr) : { d: 0, at: [] };
        dx += sx.d;
        dy += sy.d;
        ui.guides = { x: sx.at, y: sy.at };
        const map = new Map<string, Partial<El>>();
        for (const o of d.orig) map.set(o.id, moveEl(o, Math.round(dx), Math.round(dy)));
        doc.setPreview(map);
        ui.emit();
        return;
      }
      case 'resize': {
        const map = new Map<string, Partial<El>>();
        const keep = d.aspect !== e.shiftKey || (d.orig.length > 1 && e.shiftKey);
        const fromCenter = e.altKey || isMod(e);
        if (d.orig.length === 1) {
          const o = d.orig[0];
          const f = resizeFrame(d.box, d.rot, d.d, p, keep, fromCenter);
          let fr = f;
          if (thr && !d.rot) {
            const t = snapTargets(new Set(d.ids));
            fr = snapResize(f, d.d, t, thr, ui);
          } else ui.guides = null;
          map.set(o.id, setFrame(o, roundFrame(fr)));
        } else {
          const f = resizeFrame(d.box, 0, d.d, p, keep, fromCenter);
          const sx = d.box.w ? f.w / d.box.w : 1;
          const sy = d.box.h ? f.h / d.box.h : 1;
          for (const o of d.orig) map.set(o.id, setFrame(o, roundFrame({ x: f.x + (o.x - d.box.x) * sx, y: f.y + (o.y - d.box.y) * sy, w: o.w * sx, h: o.h * sy })));
        }
        doc.setPreview(map);
        ui.emit();
        return;
      }
      case 'rotate': {
        let delta = ((Math.atan2(p.y - d.cy, p.x - d.cx) - d.a0) * 180) / Math.PI;
        const map = new Map<string, Partial<El>>();
        for (const o of d.orig) {
          let r = (o.rot ?? 0) + delta;
          if (e.shiftKey) r = Math.round(r / 15) * 15;
          else if (Math.abs(((r % 90) + 90) % 90) < 3 || Math.abs(((r % 90) + 90) % 90) > 87) r = Math.round(r / 90) * 90;
          r = ((r % 360) + 360) % 360;
          delta = r - (o.rot ?? 0);
          map.set(o.id, { rot: Math.round(r * 10) / 10 });
        }
        doc.setPreview(map);
        return;
      }
      case 'adjust': {
        const o = d.orig;
        const hs = adjHandles(o.geom)[d.idx];
        if (!hs) return;
        const local = toLocal(o, p);
        const adj = { ...(DEFAULT_ADJ[o.geom] ?? {}), ...(o.adj ?? {}), ...hs.apply(o.w, o.h, local.x, local.y) };
        doc.setPreview(new Map([[o.id, { adj }]]));
        return;
      }
      case 'endpoint': {
        const o = d.orig;
        const ends = lineEnds(o);
        let { x, y } = p;
        const other = d.end === 'start' ? { x: ends.x2, y: ends.y2 } : { x: ends.x1, y: ends.y1 };
        if (e.shiftKey) {
          // constrain to 45° steps
          const ang = Math.round(Math.atan2(y - other.y, x - other.x) / (Math.PI / 4)) * (Math.PI / 4);
          const len = Math.hypot(x - other.x, y - other.y);
          x = other.x + Math.cos(ang) * len;
          y = other.y + Math.sin(ang) * len;
        }
        const n = d.end === 'start' ? fromEnds(o, x, y, ends.x2, ends.y2) : fromEnds(o, ends.x1, ends.y1, x, y);
        doc.setPreview(new Map([[o.id, { x: Math.round(n.x), y: Math.round(n.y), w: Math.round(n.w), h: Math.round(n.h), flipH: n.flipH, flipV: n.flipV }]]));
        return;
      }
      case 'marquee': {
        const r = { x: Math.min(d.sx, p.x), y: Math.min(d.sy, p.y), w: Math.abs(p.x - d.sx), h: Math.abs(p.y - d.sy) };
        ui.marquee = r;
        const inside = slide.elements
          .filter((el) => {
            if (el.hidden) return false;
            const b = rotatedBounds(el);
            return b.x >= r.x && b.y >= r.y && b.x + b.w <= r.x + r.w && b.y + b.h <= r.y + r.h;
          })
          .map((el) => el.id);
        doc.sel = { ...doc.sel, els: [...new Set([...d.base, ...inside])] };
        doc.emit();
        ui.emit();
        return;
      }
      case 'insert': {
        d.x = p.x;
        d.y = p.y;
        ui.marquee = { x: Math.min(d.sx, p.x), y: Math.min(d.sy, p.y), w: Math.abs(p.x - d.sx), h: Math.abs(p.y - d.sy) };
        ui.emit();
        return;
      }
    }
  };

  const onPointerUp = (e: React.PointerEvent) => {
    const d = dragRef.current;
    dragRef.current = null;
    wrapRef.current?.releasePointerCapture?.(e.pointerId);
    const preview = doc.preview;
    ui.guides = null;
    ui.marquee = null;
    if (!d) return;
    switch (d.kind) {
      case 'move':
        if (d.moved && preview) {
          if (d.copy) {
            doc.preview = null;
            insertEls(doc, cloneEls(d.orig.map((o) => preview.get(o.id) as El)), 'Copy');
          } else {
            doc.preview = null;
            doc.updateEls('Move', d.ids, (el) => (preview.get(el.id) as El) ?? el);
          }
        } else doc.setPreview(null);
        break;
      case 'resize':
      case 'rotate':
      case 'adjust':
      case 'endpoint':
        doc.preview = null;
        if (preview?.size) doc.updateEls(d.kind === 'rotate' ? 'Rotate' : d.kind === 'adjust' ? 'Adjust shape' : 'Resize', [...preview.keys()], (el) => ({ ...el, ...(preview.get(el.id) ?? {}) }) as El);
        else doc.emit();
        break;
      case 'marquee':
        ui.emit();
        break;
      case 'insert': {
        const mode = ui.insert!;
        ui.insert = null;
        const dragged = Math.abs(d.x - d.sx) * zoom > 4 || Math.abs(d.y - d.sy) * zoom > 4;
        let x = Math.round(Math.min(d.sx, d.x));
        let y = Math.round(Math.min(d.sy, d.y));
        let w = Math.round(Math.abs(d.x - d.sx));
        let hh = Math.round(Math.abs(d.y - d.sy));
        if (mode.kind === 'textbox') {
          if (!dragged) w = 360;
          const tb = textBox(x, y, Math.max(40, w));
          insertEls(doc, [tb], 'Text box');
          startEdit(tb.id);
        } else if (mode.kind === 'table') {
          if (!dragged) {
            w = Math.min(pres.size.w - 2 * x, 160 * mode.cols);
          }
          insertEls(doc, [createTable(mode.rows, mode.cols, x, y, Math.max(120, w))], 'Table');
        } else {
          const line = mode.geom === 'line' || mode.geom.startsWith('straightConnector') || mode.geom.startsWith('bentConnector') || mode.geom.startsWith('curvedConnector');
          if (!dragged) {
            w = line ? 200 : 160;
            hh = line ? 0 : 160;
            x -= line ? 0 : w / 2;
            y -= line ? 0 : hh / 2;
          }
          if (e.shiftKey && !line) hh = w = Math.max(w, hh);
          const s = shapeEl(mode.geom, Math.round(x), Math.round(y), Math.max(line ? 0 : 4, w), Math.max(line ? 0 : 4, hh));
          if (line && dragged) Object.assign(s, fromEnds(s, d.sx, d.sy, d.x, d.y));
          insertEls(doc, [s], 'Insert shape');
        }
        ui.emit();
        break;
      }
    }
  };

  const onDoubleClick = (e: React.MouseEvent) => {
    // pointer capture retargets click events to the wrapper, so look up what is really under the pointer
    const target = (document.elementFromPoint(e.clientX, e.clientY) as HTMLElement | null) ?? (e.target as HTMLElement);
    if (target.closest('.sl-pm')) return;
    const id = topLevelId(target);
    if (!id) return;
    const el = findEl(slide.elements, id);
    if (!el) return;
    if (el.type === 'chart') {
      h.onEditChart(id);
      return;
    }
    if (el.type === 'table') {
      const td = target.closest<HTMLElement>('td[data-cell]');
      if (td) {
        const [r, c] = td.dataset.cell!.split(':').map(Number);
        startEdit(id, [r, c], { x: e.clientX, y: e.clientY });
      }
      return;
    }
    if (el.type === 'shape' && !isLineEl(el)) startEdit(id, undefined, { x: e.clientX, y: e.clientY });
  };

  /* -------------------------------------------------------- keyboard */

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (doc.editing) return;
    if (h.onShortcut(e)) {
      e.preventDefault();
      return;
    }
    const mod = isMod(e);
    const sel = doc.sel.els;
    const k = e.key;
    const step = e.shiftKey ? 10 : 1;
    switch (k) {
      case 'ArrowLeft':
      case 'ArrowRight':
      case 'ArrowUp':
      case 'ArrowDown':
        e.preventDefault();
        if (!sel.length) {
          if (k === 'ArrowUp' || k === 'ArrowLeft') doc.goToSlide(doc.sel.slide - 1);
          else doc.goToSlide(doc.sel.slide + 1);
          return;
        }
        nudge(doc, k === 'ArrowLeft' ? -step : k === 'ArrowRight' ? step : 0, k === 'ArrowUp' ? -step : k === 'ArrowDown' ? step : 0);
        return;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        if (sel.length) deleteEls(doc);
        return;
      case 'Escape':
        e.preventDefault();
        if (ui.insert) ui.setInsert(null);
        else doc.selectEls([]);
        return;
      case 'Tab': {
        e.preventDefault();
        const els = slide.elements;
        if (!els.length) return;
        const cur = els.findIndex((x) => x.id === sel[sel.length - 1]);
        const next = (cur + (e.shiftKey ? -1 : 1) + els.length) % els.length;
        doc.selectEls([els[cur < 0 ? 0 : next].id]);
        return;
      }
      case 'Enter':
      case 'F2':
        if (sel.length === 1) {
          e.preventDefault();
          const el = findEl(slide.elements, sel[0]);
          if (el?.type === 'shape' && !isLineEl(el)) startEdit(sel[0], undefined, k === 'F2' ? 'end' : 'all');
          else if (el?.type === 'table') startEdit(sel[0], [0, 0], 'end');
        }
        return;
      case 'PageDown':
        e.preventDefault();
        doc.goToSlide(doc.sel.slide + 1);
        return;
      case 'PageUp':
        e.preventDefault();
        doc.goToSlide(doc.sel.slide - 1);
        return;
      case 'Home':
        if (!sel.length) {
          e.preventDefault();
          doc.goToSlide(0);
        }
        return;
      case 'End':
        if (!sel.length) {
          e.preventDefault();
          doc.goToSlide(pres.slides.length - 1);
        }
        return;
      case 'F5':
        e.preventDefault();
        h.onStartShow(e.shiftKey);
        return;
    }
    if (mod) {
      const low = k.toLowerCase();
      if (low === 'a') {
        e.preventDefault();
        doc.selectEls(slide.elements.filter((x) => !x.hidden).map((x) => x.id));
      } else if (low === 'd') {
        e.preventDefault();
        duplicateEls(doc);
      } else if (low === 'g' && !e.shiftKey) {
        e.preventDefault();
        group(doc);
      } else if (low === 'g' && e.shiftKey) {
        e.preventDefault();
        ungroup(doc);
      }
      return;
    }
    // typing into a selected shape replaces its text (PowerPoint behaviour)
    if (k.length === 1 && !e.altKey && sel.length === 1) {
      const el = findEl(slide.elements, sel[0]);
      if (el?.type === 'shape' && !isLineEl(el) && !el.locked) {
        e.preventDefault();
        ui.pendingText = k;
        startEdit(sel[0], undefined, 'all');
      }
    }
  };

  const onCopy = (e: React.ClipboardEvent, cut: boolean) => {
    if (doc.editing || !doc.sel.els.length) return;
    e.preventDefault();
    const text = copyEls(doc, cut);
    e.clipboardData.setData('text/plain', text || ' ');
  };

  const onPaste = (e: React.ClipboardEvent) => {
    if (doc.editing) return;
    e.preventDefault();
    const items = Array.from(e.clipboardData.items);
    const img = items.find((it) => it.kind === 'file' && it.type.startsWith('image/'));
    const text = e.clipboardData.getData('text/plain');
    const internal = clipboardText();
    if (internal !== null && (text === internal || (internal === '' && text === ' ') || (!text && !img))) {
      pasteEls(doc);
      return;
    }
    if (img) {
      const f = img.getAsFile();
      if (f) h.onPasteImage(f);
      return;
    }
    if (text) {
      const tb = textBox(Math.round(pres.size.w / 2 - 300), Math.round(pres.size.h / 2 - 40), 600, text.replace(/\r\n?/g, '\n'));
      insertEls(doc, [tb], 'Paste');
    } else pasteEls(doc);
  };

  const onWheel = (e: React.WheelEvent) => {
    if (isMod(e)) {
      e.preventDefault();
      const z = Math.max(0.1, Math.min(4, zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)));
      ui.setZoom(Math.round(z * 100) / 100);
      return;
    }
    // with the whole slide visible, the wheel moves between slides (like PowerPoint)
    const wrap = wrapRef.current;
    if (!wrap || doc.editing) return;
    if (wrap.scrollHeight > wrap.clientHeight + 2) return;
    wheelAcc.current += e.deltaY;
    if (Math.abs(wheelAcc.current) > 120) {
      doc.goToSlide(doc.sel.slide + (wheelAcc.current > 0 ? 1 : -1));
      wheelAcc.current = 0;
    }
  };

  /* ---------------------------------------------------------- render */

  const editingKey = doc.editing ? (doc.editing.cell ? `${doc.editing.el}:${doc.editing.cell[0]}:${doc.editing.cell[1]}` : doc.editing.el) : null;
  const stageW = pres.size.w * zoom;
  const stageH = pres.size.h * zoom;

  if (!slide) return <div className="sl-canvas" ref={wrapRef} />;
  return (
    <div
      className={`sl-canvas${ui.insert ? ' inserting' : ''}${ui.painter ? ' painting' : ''}`}
      ref={wrapRef}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      onPointerCancel={onPointerUp}
      onDoubleClick={onDoubleClick}
      onWheel={onWheel}
      onContextMenu={(e) => {
        if ((e.target as HTMLElement).closest('.sl-pm')) return;
        e.preventDefault();
        const id = topLevelId(e.target as HTMLElement);
        if (id && !doc.sel.els.includes(id)) doc.selectEls([id]);
        h.onContextMenu(e, id ? (findEl(slide.elements, id) ?? null) : null);
      }}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('Files')) e.preventDefault();
      }}
      onDrop={(e) => {
        const files = Array.from(e.dataTransfer.files);
        if (!files.length) return;
        e.preventDefault();
        e.stopPropagation();
        h.onDropFiles(files, toSlide(e.clientX, e.clientY));
      }}
    >
      <div className="sl-stage" style={{ width: stageW, height: stageH, margin: `${pad}px auto` }}>
        <div className="sl-stage-inner" ref={stageRef} style={{ width: pres.size.w, height: pres.size.h, transform: `scale(${zoom})` }}>
          <SlideView pres={pres} slide={slide} slideNumber={doc.sel.slide + 1} mode="edit" editing={editingKey} preview={doc.preview ?? undefined}>
            {ui.showGrid && <div className="sl-gridlines" style={{ backgroundSize: `${pres.size.w / 40}px ${pres.size.w / 40}px` }} />}
            <Overlay doc={doc} ui={ui} zoom={zoom} />
          </SlideView>
        </div>
      </div>
      {doc.editing && <TextEditSession key={editingKey!} doc={doc} ui={ui} target={doc.editing} root={() => stageRef.current} onExit={exitEdit} onTab={tabCell} />}
      <textarea
        ref={sinkRef}
        className="sl-sink"
        aria-label="Slide canvas"
        onKeyDown={onKeyDown}
        onCopy={(e) => onCopy(e, false)}
        onCut={(e) => onCopy(e, true)}
        onPaste={onPaste}
        readOnly
        value=""
      />
    </div>
  );
}

/* ================================================================ math */

function roundFrame(f: Frame): Frame {
  return { x: Math.round(f.x), y: Math.round(f.y), w: Math.max(1, Math.round(f.w)), h: Math.max(1, Math.round(f.h)) };
}

function toLocal(o: El, p: { x: number; y: number }): { x: number; y: number } {
  const cx = o.x + o.w / 2;
  const cy = o.y + o.h / 2;
  const a = (-(o.rot ?? 0) * Math.PI) / 180;
  const dx = p.x - cx;
  const dy = p.y - cy;
  let x = dx * Math.cos(a) - dy * Math.sin(a) + o.w / 2;
  const y = dx * Math.sin(a) + dy * Math.cos(a) + o.h / 2;
  if (o.flipH) x = o.w - x;
  return { x, y: o.flipV ? o.h - y : y };
}

/** Resizes a (possibly rotated) frame by dragging a handle; the opposite side stays put. */
function resizeFrame(b: Frame, rot: number, d: HandleDir, p: { x: number; y: number }, keep: boolean, fromCenter: boolean): Frame {
  const a = (rot * Math.PI) / 180;
  const cx = b.x + b.w / 2;
  const cy = b.y + b.h / 2;
  const cos = Math.cos(a);
  const sin = Math.sin(a);
  // pointer in the frame's local coordinates (origin at the centre)
  const lx = (p.x - cx) * cos + (p.y - cy) * sin;
  const ly = -(p.x - cx) * sin + (p.y - cy) * cos;
  const ax = fromCenter ? 0 : (-d[0] * b.w) / 2;
  const ay = fromCenter ? 0 : (-d[1] * b.h) / 2;
  let w = d[0] ? Math.max(2, fromCenter ? 2 * Math.abs(lx) : (lx - ax) * d[0]) : b.w;
  let h = d[1] ? Math.max(2, fromCenter ? 2 * Math.abs(ly) : (ly - ay) * d[1]) : b.h;
  if (keep && b.w > 0 && b.h > 0) {
    const s = d[0] && d[1] ? Math.max(w / b.w, h / b.h) : d[0] ? w / b.w : h / b.h;
    w = b.w * s;
    h = b.h * s;
  }
  // new centre in local coordinates
  const ncx = fromCenter ? 0 : d[0] ? ax + (d[0] * w) / 2 : 0;
  const ncy = fromCenter ? 0 : d[1] ? ay + (d[1] * h) / 2 : 0;
  const gx = cx + ncx * cos - ncy * sin;
  const gy = cy + ncx * sin + ncy * cos;
  return { x: gx - w / 2, y: gy - h / 2, w, h };
}

function snapResize(f: Frame, d: HandleDir, t: { xs: number[]; ys: number[] }, thr: number, ui: SlidesUI): Frame {
  const out = { ...f };
  const gx: number[] = [];
  const gy: number[] = [];
  const near = (v: number, cands: number[]) => {
    let best: number | null = null;
    for (const c of cands) if (Math.abs(c - v) <= thr && (best === null || Math.abs(c - v) < Math.abs(best - v))) best = c;
    return best;
  };
  if (d[0] === 1) {
    const c = near(f.x + f.w, t.xs);
    if (c !== null) {
      out.w = c - f.x;
      gx.push(c);
    }
  } else if (d[0] === -1) {
    const c = near(f.x, t.xs);
    if (c !== null) {
      out.w = f.x + f.w - c;
      out.x = c;
      gx.push(c);
    }
  }
  if (d[1] === 1) {
    const c = near(f.y + f.h, t.ys);
    if (c !== null) {
      out.h = c - f.y;
      gy.push(c);
    }
  } else if (d[1] === -1) {
    const c = near(f.y, t.ys);
    if (c !== null) {
      out.h = f.y + f.h - c;
      out.y = c;
      gy.push(c);
    }
  }
  ui.guides = gx.length || gy.length ? { x: gx, y: gy } : null;
  return out;
}

/* ============================================================= overlay */

function Overlay({ doc, ui, zoom }: { doc: SlidesDoc; ui: SlidesUI; zoom: number }) {
  const s = 1 / zoom;
  const preview = doc.preview;
  const sel = doc.selected.map((e) => (preview?.get(e.id) ? ({ ...e, ...preview.get(e.id) } as El) : e));
  const hs = 9 * s;
  const handle = (id: string, x: number, y: number, cursor: string, extra = '') => (
    <div key={id} className={`sl-handle ${extra}`} data-handle={id} style={{ left: x - hs / 2, top: y - hs / 2, width: hs, height: hs, cursor, borderWidth: s * 1.5 }} />
  );
  const editing = doc.editing;
  const content: React.ReactNode[] = [];
  if (sel.length === 1 && !editing) {
    const e = sel[0];
    const frame: CSSProperties = { left: e.x, top: e.y, width: e.w, height: e.h, transform: e.rot ? `rotate(${e.rot}deg)` : undefined, borderWidth: s * 1.5 };
    if (isLineEl(e)) {
      const ends = lineEnds(e as ShapeEl);
      content.push(
        <div key="line" className="sl-sel-line">
          {handle('start', ends.x1, ends.y1, 'move', 'round')}
          {handle('end', ends.x2, ends.y2, 'move', 'round')}
        </div>,
      );
    } else {
      const adj = e.type === 'shape' && !e.locked ? adjHandles(e.geom) : [];
      content.push(
        <div key="box" className={`sl-sel-box${e.locked ? ' locked' : ''}`} style={frame}>
          {!e.locked && (
            <>
              <div className="sl-rot-stem" style={{ height: 22 * s, width: s, top: -22 * s }} />
              {handle('rotate', e.w / 2, -24 * s, 'grab', 'round rotate')}
              {HANDLES.map((hd) => handle(hd.id, ((hd.d[0] + 1) / 2) * e.w, ((hd.d[1] + 1) / 2) * e.h, cursorFor(hd.d, e.rot ?? 0)))}
              {adj.map((a, i) => {
                const sh = e as ShapeEl;
                let [ax, ay] = a.pos(sh.w, sh.h, { ...(DEFAULT_ADJ[sh.geom] ?? {}), ...(sh.adj ?? {}) });
                if (sh.flipH) ax = sh.w - ax;
                if (sh.flipV) ay = sh.h - ay;
                return handle(`adj${i}`, ax, ay, 'pointer', 'adj');
              })}
            </>
          )}
        </div>,
      );
    }
  } else if (sel.length > 1) {
    for (const e of sel) content.push(<div key={e.id} className="sl-sel-box thin" style={{ left: e.x, top: e.y, width: e.w, height: e.h, transform: e.rot ? `rotate(${e.rot}deg)` : undefined, borderWidth: s }} />);
    const b = unionBounds(sel.map((e) => ({ x: e.x, y: e.y, w: e.w, h: e.h })))!;
    content.push(
      <div key="multi" className="sl-sel-box multi" style={{ left: b.x, top: b.y, width: b.w, height: b.h, borderWidth: s * 1.5 }}>
        <div className="sl-rot-stem" style={{ height: 22 * s, width: s, top: -22 * s }} />
        {handle('rotate', b.w / 2, -24 * s, 'grab', 'round rotate')}
        {HANDLES.map((hd) => handle(hd.id, ((hd.d[0] + 1) / 2) * b.w, ((hd.d[1] + 1) / 2) * b.h, cursorFor(hd.d, 0)))}
      </div>,
    );
  } else if (editing) {
    const e = findEl(doc.slide.elements, editing.el);
    if (e) content.push(<div key="edit" className="sl-sel-box editing" style={{ left: e.x, top: e.y, width: e.w, height: e.h, transform: e.rot ? `rotate(${e.rot}deg)` : undefined, borderWidth: s * 1.5 }} />);
  }
  const g = ui.guides;
  return (
    <div className="sl-overlay">
      {content}
      {g?.x.map((x, i) => <div key={`gx${i}`} className="sl-guide v" style={{ left: x - s / 2, width: s }} />)}
      {g?.y.map((y, i) => <div key={`gy${i}`} className="sl-guide h" style={{ top: y - s / 2, height: s }} />)}
      {ui.marquee && <div className={`sl-marquee${ui.insert ? ' insert' : ''}`} style={{ left: ui.marquee.x, top: ui.marquee.y, width: ui.marquee.w, height: ui.marquee.h, borderWidth: s }} />}
    </div>
  );
}
