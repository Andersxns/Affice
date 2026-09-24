import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import { isMod } from '@/lib/utils';
import { MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import type { CellPos, SheetDoc } from '../doc';
import type { SheetUI } from '../ui/controller';
import { analyzeFormula, applySuggestion, FormulaAssist } from '../ui/formulaAssist';
import { CfEvaluator } from './cf';
import { geometryOf } from './geometry';
import { cellFont, clearMeasureCache, headerWidth, Mapper, renderGrid, type GridTheme, type Viewport } from './render';
import { ObjectLayer } from './ObjectLayer';

export type GridTarget =
  | { kind: 'cell'; r: number; c: number }
  | { kind: 'row'; r: number }
  | { kind: 'col'; c: number }
  | { kind: 'corner' };

export interface GridHandlers {
  onContextMenu: (x: number, y: number, target: GridTarget) => void;
  onFilterMenu: (col: number, anchor: DOMRect) => void;
  onListPick: (anchor: DOMRect) => void;
  onOpenLink: (url: string) => void;
  onShortcut: (e: KeyboardEvent | React.KeyboardEvent) => boolean;
  onCopy: (e: ClipboardEvent, cut: boolean) => void;
  onPaste: (e: ClipboardEvent) => void;
  onResize: (axis: 'row' | 'col', idx: number[], size: number) => void;
  onAutoFit: (axis: 'row' | 'col', idx: number[]) => void;
  onFill: (src: Range, target: Range, copyOnly: boolean) => void;
  onMove: (src: Range, dr: number, dc: number, copy: boolean) => void;
  onCommitted?: () => void;
  onSelectionEnd?: () => void;
}

type Drag =
  | { kind: 'select'; anchor: CellPos; add: boolean; header?: 'row' | 'col' }
  | { kind: 'point'; anchor: CellPos }
  | { kind: 'resize'; axis: 'row' | 'col'; idx: number; start: number; startSize: number; targets: number[]; orig: [number, unknown][] }
  | { kind: 'fill'; src: Range; target: Range }
  | { kind: 'move'; src: Range; grab: CellPos; dest: Range; copy: boolean };

function readTheme(el: HTMLElement): GridTheme {
  const cs = getComputedStyle(el);
  const v = (n: string, d: string) => cs.getPropertyValue(n).trim() || d;
  return {
    bg: v('--grid-bg', '#ffffff'),
    line: v('--grid-line', '#e1e4ea'),
    text: v('--grid-text', '#1a1d24'),
    headerBg: v('--grid-header-bg', '#f6f7f9'),
    headerFg: v('--grid-header-fg', '#5c6370'),
    headerLine: v('--grid-header-line', '#d9dce3'),
    headerSelBg: v('--grid-header-sel', '#e3f3ea'),
    headerSelFg: v('--grid-header-sel-fg', '#0a6b38'),
    accent: v('--grid-accent', '#17a35a'),
    selFill: v('--grid-sel-fill', 'rgba(23,163,90,0.12)'),
    frozenLine: v('--grid-frozen', '#b8bec9'),
    uiFont: v('--font-ui', 'system-ui, sans-serif'),
    dark: document.documentElement.getAttribute('data-theme') === 'dark',
  };
}

export function Grid({ doc, ui, active, h }: { doc: SheetDoc; ui: SheetUI; active: boolean; h: GridHandlers }) {
  const docVersion = useSyncExternalStore(doc.subscribe, doc.getVersion);
  const uiVersion = useSyncExternalStore(ui.subscribe, ui.getVersion);
  const rootRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const scrollerRef = useRef<HTMLDivElement>(null);
  const spacerRef = useRef<HTMLDivElement>(null);
  const sinkRef = useRef<HTMLTextAreaElement>(null);
  const editorRef = useRef<HTMLTextAreaElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });
  const [theme, setTheme] = useState<GridTheme | null>(null);
  const [cursor, setCursor] = useState('cell');
  const [hoverNote, setHoverNote] = useState<{ x: number; y: number; text: string } | null>(null);
  const [assistIndex, setAssistIndex] = useState(0);
  const dragRef = useRef<Drag | null>(null);
  const [dragTick, setDragTick] = useState(0);
  const mapperRef = useRef<Mapper | null>(null);
  const cfRef = useRef<CfEvaluator>(new CfEvaluator(doc));
  const frame = useRef(0);
  const syncingScroll = useRef(false);
  const lastPointer = useRef<{ x: number; y: number } | null>(null);
  const autoScroll = useRef<number | null>(null);
  const composing = useRef(false);

  const sheet = doc.sheet;
  const zoom = ui.zoom;
  const edit = ui.edit;

  /* ------------------------------------------------------------ theme */
  useEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    const apply = () => {
      setTheme(readTheme(el));
      clearMeasureCache();
    };
    apply();
    const mo = new MutationObserver(apply);
    mo.observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme', 'style', 'class'] });
    const fontsDone = () => {
      clearMeasureCache();
      requestDraw();
    };
    document.fonts?.addEventListener?.('loadingdone', fontsDone);
    return () => {
      mo.disconnect();
      document.fonts?.removeEventListener?.('loadingdone', fontsDone);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* ------------------------------------------------------------- size */
  useLayoutEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  /* -------------------------------------------------------- viewport */
  const ext = doc.engine.extent(sheet.id);
  const headerW = sheet.view.showHeaders ? headerWidth(Math.max(ext.rows, ui.getReach().rows, 999), zoom) : 0;
  const headerH = sheet.view.showHeaders ? Math.round(24 * Math.min(1.3, Math.max(0.85, zoom))) : 0;
  const geo = geometryOf(sheet, docVersion);
  const fr = Math.min(sheet.view.freezeRows, MAX_ROWS - 1);
  const fc = Math.min(sheet.view.freezeCols, MAX_COLS - 1);
  ui.viewport = { width: size.w, height: size.h, headerW, headerH };

  const viewport = useCallback((): Viewport => {
    const s = ui.getScroll();
    return { width: size.w, height: size.h, headerW, headerH, zoom, scrollX: s.x, scrollY: s.y, fr, fc };
  }, [size.w, size.h, headerW, headerH, zoom, fr, fc, ui]);

  // virtual content size for the native scrollbars
  const contentSize = useCallback(() => {
    const reach = ui.getReach();
    const rows = Math.min(MAX_ROWS, Math.max(ext.rows + 60, reach.rows, 100, fr + 1));
    const cols = Math.min(MAX_COLS, Math.max(ext.cols + 12, reach.cols, 30, fc + 1));
    const s = ui.getScroll();
    const w = (geo.cols.pos(cols) - geo.cols.pos(fc)) * zoom + headerW + geo.cols.pos(fc) * zoom;
    const hh = (geo.rows.pos(rows) - geo.rows.pos(fr)) * zoom + headerH + geo.rows.pos(fr) * zoom;
    return { w: Math.max(w, s.x * zoom + size.w + 200), h: Math.max(hh, s.y * zoom + size.h + 200) };
  }, [ui, ext.rows, ext.cols, fr, fc, geo, zoom, headerW, headerH, size.w, size.h]);

  /* ------------------------------------------------------------- draw */
  const draw = useCallback(() => {
    frame.current = 0;
    const canvas = canvasRef.current;
    if (!canvas || !theme || size.w <= 0 || size.h <= 0) return;
    const dpr = window.devicePixelRatio || 1;
    const W = Math.round(size.w * dpr);
    const H = Math.round(size.h * dpr);
    if (canvas.width !== W || canvas.height !== H) {
      canvas.width = W;
      canvas.height = H;
    }
    const ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) return;
    const vp = viewport();
    const g = geometryOf(doc.sheet, doc.version);
    const m = new Mapper(g, vp);
    mapperRef.current = m;
    const d = dragRef.current;
    const e = ui.edit;
    const refs = e && e.sheet === doc.sheet.id ? ui.editRefs() : e ? ui.editRefs() : [];
    const activeSpill = doc.engine.spillAreaOf(doc.sheet.id, doc.sel.active.r, doc.sel.active.c) ?? (() => {
      const a = doc.engine.spillAnchorOf(doc.sheet.id, doc.sel.active.r, doc.sel.active.c);
      return a ? doc.engine.spillAreaOf(doc.sheet.id, a.r, a.c) : undefined;
    })();
    renderGrid({
      ctx,
      doc,
      sheet: doc.sheet,
      geo: g,
      vp,
      dpr,
      theme,
      cf: cfRef.current,
      ov: {
        sel: doc.sel,
        editing: e && e.sheet === doc.sheet.id ? { r: e.r, c: e.c } : null,
        refs,
        clip: doc.clip && doc.clip.sheet === doc.sheet.id ? { range: doc.clip.range, phase: ui.clipPhase } : null,
        fillPreview: d?.kind === 'fill' ? d.target : null,
        movePreview: d?.kind === 'move' ? d.dest : null,
        showFormulas: ui.showFormulas,
        focused: active,
        findHits: ui.findHits ?? undefined,
        spillHint: activeSpill ? { r1: activeSpill.r1, c1: activeSpill.c1, r2: activeSpill.r2, c2: activeSpill.c2 } : null,
      },
    });
  }, [theme, size.w, size.h, viewport, doc, ui, active]);

  const requestDraw = useCallback(() => {
    if (!frame.current) frame.current = requestAnimationFrame(draw);
  }, [draw]);

  useEffect(() => {
    requestDraw();
  }, [requestDraw, docVersion, uiVersion, dragTick]);

  useEffect(
    () => () => {
      cancelAnimationFrame(frame.current);
      frame.current = 0;
    },
    [],
  );

  // marching ants for the copy source
  useEffect(() => {
    if (!doc.clip || !active) return;
    const t = setInterval(() => {
      ui.clipPhase = (ui.clipPhase + 1) % 18;
      requestDraw();
    }, 60);
    return () => clearInterval(t);
  }, [doc.clip, active, ui, requestDraw, docVersion]);

  /* ---------------------------------------------------------- scroll */
  // keep the native scroller in sync with the model scroll (keyboard navigation etc.)
  const syncScroller = useCallback(() => {
    const sc = scrollerRef.current;
    const sp = spacerRef.current;
    if (!sc || !sp) return;
    const cs = contentSize();
    sp.style.width = `${cs.w}px`;
    sp.style.height = `${cs.h}px`;
    const s = ui.getScroll();
    const x = Math.round(s.x * zoom);
    const y = Math.round(s.y * zoom);
    if (Math.abs(sc.scrollLeft - x) > 1 || Math.abs(sc.scrollTop - y) > 1) {
      syncingScroll.current = true;
      sc.scrollLeft = x;
      sc.scrollTop = y;
      requestAnimationFrame(() => (syncingScroll.current = false));
    }
  }, [contentSize, ui, zoom]);

  useLayoutEffect(() => {
    syncScroller();
  });

  useEffect(() => {
    ui.onScrollRequest = () => {
      syncScroller();
      requestDraw();
    };
    return () => {
      ui.onScrollRequest = undefined;
    };
  }, [ui, syncScroller, requestDraw]);

  const onScroll = () => {
    const sc = scrollerRef.current;
    if (!sc || syncingScroll.current) return;
    ui.setScroll(sc.scrollLeft / zoom, sc.scrollTop / zoom);
    // grow the scrollable area as the user approaches the end
    const reach = ui.getReach();
    const bottomRow = geo.rows.at(geo.rows.pos(fr) + sc.scrollTop / zoom + size.h / zoom);
    const rightCol = geo.cols.at(geo.cols.pos(fc) + sc.scrollLeft / zoom + size.w / zoom);
    if (bottomRow + 40 > Math.max(reach.rows, ext.rows + 60)) reach.rows = Math.min(MAX_ROWS, bottomRow + 200);
    if (rightCol + 10 > Math.max(reach.cols, ext.cols + 12)) reach.cols = Math.min(MAX_COLS, rightCol + 30);
    const sp = spacerRef.current;
    if (sp) {
      const cs = contentSize();
      sp.style.width = `${cs.w}px`;
      sp.style.height = `${cs.h}px`;
    }
    requestDraw();
    setHoverNote(null);
  };

  /* ------------------------------------------------------- focus */
  const focusSink = useCallback(() => {
    const s = sinkRef.current;
    if (!s) return;
    s.value = ' ';
    s.select();
    s.focus({ preventScroll: true });
  }, []);

  useEffect(() => {
    if (!active) return;
    if (ui.edit && ui.edit.source === 'cell') editorRef.current?.focus({ preventScroll: true });
    else if (!ui.edit && document.activeElement?.closest('.formula-bar') == null && !document.activeElement?.closest('.dialog, .popover, .menu, input, textarea:not(.grid-sink)')) focusSink();
  }, [active, uiVersion, focusSink, ui]);

  /* ----------------------------------------------------- hit testing */
  const local = (e: { clientX: number; clientY: number }) => {
    const r = rootRef.current!.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };

  const hit = useCallback(
    (x: number, y: number) => {
      const m = mapperRef.current;
      if (!m) return { zone: 'none' as const };
      const inHeaderRow = y < headerH;
      const inHeaderCol = x < headerW;
      if (inHeaderRow && inHeaderCol) return { zone: 'corner' as const };
      if (inHeaderRow) {
        const c = m.colAt(x);
        const left = m.x(c);
        const right = m.x(c + 1);
        if (Math.abs(x - right) <= 4) return { zone: 'colResize' as const, c };
        if (Math.abs(x - left) <= 3 && c > 0) return { zone: 'colResize' as const, c: geo.cols.step(c, -1) === c ? c - 1 : c - 1 };
        return { zone: 'colHeader' as const, c };
      }
      if (inHeaderCol) {
        const r = m.rowAt(y);
        const top = m.y(r);
        const bottom = m.y(r + 1);
        if (Math.abs(y - bottom) <= 3) return { zone: 'rowResize' as const, r };
        if (Math.abs(y - top) <= 2 && r > 0) return { zone: 'rowResize' as const, r: r - 1 };
        return { zone: 'rowHeader' as const, r };
      }
      const pos = m.cellAt(x, y);
      const sel = doc.sel;
      if (sel.ranges.length === 1 && !ui.edit) {
        const rg = ui.withMerges(sel.ranges[0]);
        const rect = m.rect(rg);
        const hx = rect.x + rect.w;
        const hy = rect.y + rect.h;
        if (Math.abs(x - hx) <= 6 && Math.abs(y - hy) <= 6) return { zone: 'fill' as const, ...pos };
        const onEdge =
          x >= rect.x - 3 &&
          x <= hx + 3 &&
          y >= rect.y - 3 &&
          y <= hy + 3 &&
          (Math.abs(x - rect.x) <= 3 || Math.abs(x - hx) <= 3 || Math.abs(y - rect.y) <= 3 || Math.abs(y - hy) <= 3);
        if (onEdge && !(rg.r1 === 0 && rg.r2 === MAX_ROWS - 1) && !(rg.c1 === 0 && rg.c2 === MAX_COLS - 1)) return { zone: 'edge' as const, ...pos };
      }
      // filter buttons
      const af = doc.sheet.filter;
      if (af && pos.r === af.range.r1 && pos.c >= af.range.c1 && pos.c <= af.range.c2) {
        const bx = m.x(pos.c + 1) - 17 * zoom;
        if (x >= bx) return { zone: 'filter' as const, ...pos };
      }
      // validation dropdown
      const a = sel.active;
      const act = m.rect(ui.withMerges({ r1: a.r, c1: a.c, r2: a.r, c2: a.c }));
      if (x > act.x + act.w && x <= act.x + act.w + 20 * zoom && y >= act.y && y <= act.y + act.h) {
        const v = doc.sheet.validations.find((vv) => vv.type === 'list' && vv.showDropdown !== false && vv.ranges.some((rg) => a.r >= rg.r1 && a.r <= rg.r2 && a.c >= rg.c1 && a.c <= rg.c2));
        if (v) return { zone: 'list' as const, ...a };
      }
      return { zone: 'cell' as const, ...pos };
    },
    [headerH, headerW, geo, doc, ui, zoom],
  );

  /* ---------------------------------------------------- mouse down */
  const commitIfEditing = useCallback((): boolean => {
    if (!ui.edit) return true;
    const ok = ui.commitEdit();
    if (ok) h.onCommitted?.();
    return ok;
  }, [ui, h]);

  const onPointerDown = (e: React.PointerEvent) => {
    if (e.button === 1) return;
    const { x, y } = local(e);
    const t = hit(x, y);
    const sc = scrollerRef.current!;
    // ignore clicks on the native scrollbars
    if (x > sc.clientWidth || y > sc.clientHeight) return;
    const m = mapperRef.current;
    if (!m) return;
    setHoverNote(null);
    if (e.button === 2) {
      // right-click keeps a selection that contains the clicked cell
      if (t.zone === 'cell' || t.zone === 'fill' || t.zone === 'edge') {
        const inside = doc.sel.ranges.some((rg) => t.r >= rg.r1 && t.r <= rg.r2 && t.c >= rg.c1 && t.c <= rg.c2);
        if (!inside) {
          if (!commitIfEditing()) return;
          ui.selectCell(t.r, t.c);
        }
      } else if (t.zone === 'colHeader') {
        if (!doc.sel.ranges.some((rg) => rg.r1 === 0 && rg.r2 === MAX_ROWS - 1 && t.c >= rg.c1 && t.c <= rg.c2)) ui.select({ r1: 0, c1: t.c, r2: MAX_ROWS - 1, c2: t.c }, { r: m.rowAt(headerH + 1), c: t.c });
      } else if (t.zone === 'rowHeader') {
        if (!doc.sel.ranges.some((rg) => rg.c1 === 0 && rg.c2 === MAX_COLS - 1 && t.r >= rg.r1 && t.r <= rg.r2)) ui.select({ r1: t.r, c1: 0, r2: t.r, c2: MAX_COLS - 1 }, { r: t.r, c: m.colAt(headerW + 1) });
      }
      return;
    }
    // point mode: clicking cells while typing a formula inserts references
    if (ui.edit && ui.canPoint() && (t.zone === 'cell' || t.zone === 'fill' || t.zone === 'edge')) {
      e.preventDefault();
      const rg = { r1: t.r, c1: t.c, r2: t.r, c2: t.c };
      ui.pointAt(ui.withMerges(rg), { r: t.r, c: t.c }, { r: t.r, c: t.c });
      dragRef.current = { kind: 'point', anchor: { r: t.r, c: t.c } };
      sc.setPointerCapture(e.pointerId);
      return;
    }
    if (t.zone === 'colResize' || t.zone === 'rowResize') {
      e.preventDefault();
      const axis = t.zone === 'colResize' ? 'col' : 'row';
      const idx = axis === 'col' ? (t as { c: number }).c : (t as { r: number }).r;
      const selected = axis === 'col' ? doc.sel.ranges.filter((rg) => rg.r1 === 0 && rg.r2 === MAX_ROWS - 1 && idx >= rg.c1 && idx <= rg.c2) : doc.sel.ranges.filter((rg) => rg.c1 === 0 && rg.c2 === MAX_COLS - 1 && idx >= rg.r1 && idx <= rg.r2);
      const targets: number[] = [];
      if (selected.length) for (const rg of doc.sel.ranges) for (let i = axis === 'col' ? rg.c1 : rg.r1; i <= (axis === 'col' ? rg.c2 : rg.r2) && targets.length < 5000; i++) targets.push(i);
      else targets.push(idx);
      const infoMap = (axis === 'col' ? doc.sheet.cols : doc.sheet.rows) as Map<number, unknown>;
      const orig = targets.map((i) => [i, infoMap.get(i)] as [number, unknown]);
      dragRef.current = { kind: 'resize', axis, idx, start: axis === 'col' ? x : y, startSize: axis === 'col' ? geo.cols.sizeOf(idx) : geo.rows.sizeOf(idx), targets, orig };
      sc.setPointerCapture(e.pointerId);
      setDragTick((v) => v + 1);
      return;
    }
    if (!commitIfEditing()) return;
    if (t.zone === 'filter') {
      const rect = rootRef.current!.getBoundingClientRect();
      const bx = m.x(t.c + 1) - 17 * zoom;
      const by = m.y(t.r);
      h.onFilterMenu(t.c, new DOMRect(rect.left + bx, rect.top + by, 17 * zoom, m.y(t.r + 1) - by));
      return;
    }
    if (t.zone === 'list') {
      const rect = rootRef.current!.getBoundingClientRect();
      const act = m.rect(ui.withMerges({ r1: t.r, c1: t.c, r2: t.r, c2: t.c }));
      h.onListPick(new DOMRect(rect.left + act.x, rect.top + act.y, act.w, act.h));
      return;
    }
    if (t.zone === 'fill') {
      e.preventDefault();
      const src = ui.withMerges(doc.sel.ranges[0]);
      dragRef.current = { kind: 'fill', src, target: src };
      sc.setPointerCapture(e.pointerId);
      return;
    }
    if (t.zone === 'edge') {
      e.preventDefault();
      const src = ui.withMerges(doc.sel.ranges[0]);
      dragRef.current = { kind: 'move', src, grab: { r: t.r, c: t.c }, dest: src, copy: isMod(e) };
      sc.setPointerCapture(e.pointerId);
      return;
    }
    if (t.zone === 'corner') {
      ui.select({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, doc.sel.active);
      return;
    }
    const add = isMod(e);
    if (t.zone === 'colHeader') {
      const r0 = m.rowAt(headerH + 1);
      if (e.shiftKey) {
        const an = doc.sel.anchor;
        ui.select({ r1: 0, c1: Math.min(an.c, t.c), r2: MAX_ROWS - 1, c2: Math.max(an.c, t.c) }, doc.sel.active, an);
      } else ui.select({ r1: 0, c1: t.c, r2: MAX_ROWS - 1, c2: t.c }, { r: r0, c: t.c }, { r: r0, c: t.c }, add);
      dragRef.current = { kind: 'select', anchor: { r: r0, c: e.shiftKey ? doc.sel.anchor.c : t.c }, add, header: 'col' };
      sc.setPointerCapture(e.pointerId);
      return;
    }
    if (t.zone === 'rowHeader') {
      const c0 = m.colAt(headerW + 1);
      if (e.shiftKey) {
        const an = doc.sel.anchor;
        ui.select({ r1: Math.min(an.r, t.r), c1: 0, r2: Math.max(an.r, t.r), c2: MAX_COLS - 1 }, doc.sel.active, an);
      } else ui.select({ r1: t.r, c1: 0, r2: t.r, c2: MAX_COLS - 1 }, { r: t.r, c: c0 }, { r: t.r, c: c0 }, add);
      dragRef.current = { kind: 'select', anchor: { r: e.shiftKey ? doc.sel.anchor.r : t.r, c: c0 }, add, header: 'row' };
      sc.setPointerCapture(e.pointerId);
      return;
    }
    if (t.zone === 'cell') {
      // Ctrl+click opens hyperlinks
      const cell = doc.sheet.get(t.r, t.c);
      if (isMod(e) && cell?.link) {
        h.onOpenLink(cell.link);
        return;
      }
      if (e.shiftKey) ui.selectCell(t.r, t.c, true);
      else {
        const mg = doc.sheet.mergeAt(t.r, t.c);
        const rg = mg ?? { r1: t.r, c1: t.c, r2: t.r, c2: t.c };
        ui.select(rg, { r: rg.r1, c: rg.c1 }, { r: rg.r1, c: rg.c1 }, add);
      }
      dragRef.current = { kind: 'select', anchor: e.shiftKey ? doc.sel.anchor : { r: t.r, c: t.c }, add: add || e.shiftKey };
      sc.setPointerCapture(e.pointerId);
      focusSink();
    }
  };

  /* ---------------------------------------------------- drag update */
  const updateDrag = useCallback(
    (x: number, y: number) => {
      const d = dragRef.current;
      const m = mapperRef.current;
      if (!d || !m) return;
      const pos = m.cellAt(Math.max(headerW + 1, x), Math.max(headerH + 1, y));
      if (d.kind === 'select') {
        const sel = doc.sel;
        let rg: Range;
        if (d.header === 'col') rg = { r1: 0, c1: Math.min(d.anchor.c, pos.c), r2: MAX_ROWS - 1, c2: Math.max(d.anchor.c, pos.c) };
        else if (d.header === 'row') rg = { r1: Math.min(d.anchor.r, pos.r), c1: 0, r2: Math.max(d.anchor.r, pos.r), c2: MAX_COLS - 1 };
        else rg = ui.withMerges({ r1: Math.min(d.anchor.r, pos.r), c1: Math.min(d.anchor.c, pos.c), r2: Math.max(d.anchor.r, pos.r), c2: Math.max(d.anchor.c, pos.c) });
        const ranges = d.add ? [...sel.ranges.slice(0, -1), rg] : [rg];
        doc.sel = { ranges, active: sel.active, anchor: sel.anchor };
        ui.emit();
      } else if (d.kind === 'point') {
        const rg = { r1: Math.min(d.anchor.r, pos.r), c1: Math.min(d.anchor.c, pos.c), r2: Math.max(d.anchor.r, pos.r), c2: Math.max(d.anchor.c, pos.c) };
        ui.pointAt(ui.withMerges(rg), d.anchor, pos);
      } else if (d.kind === 'resize') {
        const delta = ((d.axis === 'col' ? x : y) - d.start) / zoom;
        const size = Math.max(0, Math.round(d.startSize + delta));
        const map = d.axis === 'col' ? doc.sheet.cols : doc.sheet.rows;
        for (const i of d.targets) map.set(i, { ...(map.get(i) ?? {}), ...(d.axis === 'col' ? { w: size } : { h: size }) });
        doc.version++;
        setDragTick((v) => v + 1);
      } else if (d.kind === 'fill') {
        const s = d.src;
        // extend along the axis with the larger distance from the source block
        const dr = pos.r > s.r2 ? pos.r - s.r2 : pos.r < s.r1 ? pos.r - s.r1 : 0;
        const dc = pos.c > s.c2 ? pos.c - s.c2 : pos.c < s.c1 ? pos.c - s.c1 : 0;
        let target: Range;
        if (Math.abs(dr) >= Math.abs(dc) && dr !== 0) target = dr > 0 ? { ...s, r2: pos.r } : { ...s, r1: pos.r };
        else if (dc !== 0) target = dc > 0 ? { ...s, c2: pos.c } : { ...s, c1: pos.c };
        else if (pos.r >= s.r1 && pos.r < s.r2 && s.r2 > s.r1) target = { ...s, r2: pos.r };
        else if (pos.c >= s.c1 && pos.c < s.c2 && s.c2 > s.c1) target = { ...s, c2: pos.c };
        else target = s;
        d.target = target;
        setDragTick((v) => v + 1);
      } else if (d.kind === 'move') {
        const dr = pos.r - d.grab.r;
        const dc = pos.c - d.grab.c;
        const h2 = d.src.r2 - d.src.r1;
        const w2 = d.src.c2 - d.src.c1;
        const r1 = Math.max(0, Math.min(MAX_ROWS - 1 - h2, d.src.r1 + dr));
        const c1 = Math.max(0, Math.min(MAX_COLS - 1 - w2, d.src.c1 + dc));
        d.dest = { r1, c1, r2: r1 + h2, c2: c1 + w2 };
        setDragTick((v) => v + 1);
      }
    },
    [doc, ui, zoom, headerW, headerH],
  );

  const stopAutoScroll = () => {
    if (autoScroll.current) cancelAnimationFrame(autoScroll.current);
    autoScroll.current = null;
  };

  const runAutoScroll = useCallback(() => {
    const p = lastPointer.current;
    const sc = scrollerRef.current;
    const m = mapperRef.current;
    if (!p || !sc || !m || !dragRef.current || dragRef.current.kind === 'resize') {
      autoScroll.current = null;
      return;
    }
    let dx = 0;
    let dy = 0;
    if (p.x > sc.clientWidth - 8) dx = Math.min(60, (p.x - sc.clientWidth + 8) / 2);
    else if (p.x < m.bodyX0 + 4) dx = -Math.min(60, (m.bodyX0 + 4 - p.x) / 2);
    if (p.y > sc.clientHeight - 8) dy = Math.min(60, (p.y - sc.clientHeight + 8) / 2);
    else if (p.y < m.bodyY0 + 4) dy = -Math.min(60, (m.bodyY0 + 4 - p.y) / 2);
    if (dx || dy) {
      if (dy > 0) {
        const reach = ui.getReach();
        reach.rows = Math.min(MAX_ROWS, Math.max(reach.rows, m.rowAt(sc.clientHeight) + 20));
      }
      if (dx > 0) {
        const reach = ui.getReach();
        reach.cols = Math.min(MAX_COLS, Math.max(reach.cols, m.colAt(sc.clientWidth) + 5));
      }
      syncScroller();
      sc.scrollLeft += dx;
      sc.scrollTop += dy;
      ui.setScroll(sc.scrollLeft / zoom, sc.scrollTop / zoom);
      updateDrag(p.x, p.y);
      requestDraw();
      autoScroll.current = requestAnimationFrame(runAutoScroll);
    } else autoScroll.current = null;
  }, [ui, zoom, updateDrag, requestDraw, syncScroller]);

  const onPointerMove = (e: React.PointerEvent) => {
    const { x, y } = local(e);
    lastPointer.current = { x, y };
    if (dragRef.current) {
      updateDrag(x, y);
      if (!autoScroll.current) autoScroll.current = requestAnimationFrame(runAutoScroll);
      return;
    }
    const t = hit(x, y);
    const cur =
      t.zone === 'colResize'
        ? 'col-resize'
        : t.zone === 'rowResize'
          ? 'row-resize'
          : t.zone === 'fill'
            ? 'crosshair'
            : t.zone === 'edge'
              ? 'move'
              : t.zone === 'colHeader'
                ? 'col-header'
                : t.zone === 'rowHeader'
                  ? 'row-header'
                  : t.zone === 'filter' || t.zone === 'list'
                    ? 'pointer'
                    : ui.edit && ui.canPoint()
                      ? 'point'
                      : 'cell';
    if (cur !== cursor) setCursor(cur);
    // note tooltips
    if (t.zone === 'cell') {
      const cell = doc.sheet.get(t.r, t.c);
      if (cell?.note) {
        const m = mapperRef.current!;
        const rect = m.rect({ r1: t.r, c1: t.c, r2: t.r, c2: t.c });
        if (!hoverNote || hoverNote.text !== cell.note) setHoverNote({ x: rect.x + rect.w + 6, y: rect.y, text: cell.note });
      } else if (hoverNote) setHoverNote(null);
    } else if (hoverNote) setHoverNote(null);
  };

  const onPointerUp = (e: React.PointerEvent) => {
    stopAutoScroll();
    const d = dragRef.current;
    dragRef.current = null;
    scrollerRef.current?.releasePointerCapture?.(e.pointerId);
    if (!d) return;
    if (d.kind === 'select') h.onSelectionEnd?.();
    if (d.kind === 'resize') {
      const size = (d.axis === 'col' ? doc.sheet.cols.get(d.idx)?.w : doc.sheet.rows.get(d.idx)?.h) ?? d.startSize;
      // restore the live-preview sizes, then record the change in one undoable step
      const map = (d.axis === 'col' ? doc.sheet.cols : doc.sheet.rows) as Map<number, unknown>;
      for (const [i, info] of d.orig) {
        if (info === undefined) map.delete(i);
        else map.set(i, info);
      }
      doc.version++;
      if (size !== d.startSize) h.onResize(d.axis, d.targets, size);
    } else if (d.kind === 'fill') {
      if (d.target.r1 !== d.src.r1 || d.target.r2 !== d.src.r2 || d.target.c1 !== d.src.c1 || d.target.c2 !== d.src.c2) h.onFill(d.src, d.target, isMod(e));
    } else if (d.kind === 'move') {
      const dr = d.dest.r1 - d.src.r1;
      const dc = d.dest.c1 - d.src.c1;
      if (dr || dc) h.onMove(d.src, dr, dc, isMod(e));
    }
    setDragTick((v) => v + 1);
    if (!ui.edit) focusSink();
    else editorRef.current?.focus();
  };

  const onDoubleClick = (e: React.MouseEvent) => {
    const { x, y } = local(e);
    const t = hit(x, y);
    if (t.zone === 'colResize') h.onAutoFit('col', selectedFull('col', t.c));
    else if (t.zone === 'rowResize') h.onAutoFit('row', selectedFull('row', t.r));
    else if (t.zone === 'fill') {
      // double-click fill handle: fill down alongside the adjacent data
      const src = ui.withMerges(doc.sel.ranges[0]);
      const col = src.c1 > 0 ? src.c1 - 1 : src.c2 + 1;
      let r = src.r2;
      const filled = (rr: number, cc: number) => {
        const v = doc.value(doc.sheet, rr, cc);
        return v !== undefined && v !== null && v !== '';
      };
      while (r + 1 < MAX_ROWS && filled(r + 1, col)) r++;
      if (r > src.r2) h.onFill(src, { ...src, r2: r }, false);
    } else if (t.zone === 'cell') {
      if (ui.edit) return;
      ui.selectCell(t.r, t.c);
      ui.startEdit(null, 'edit');
    }
  };

  const selectedFull = (axis: 'row' | 'col', idx: number): number[] => {
    const sel = doc.sel.ranges.filter((rg) => (axis === 'col' ? rg.r1 === 0 && rg.r2 === MAX_ROWS - 1 && idx >= rg.c1 && idx <= rg.c2 : rg.c1 === 0 && rg.c2 === MAX_COLS - 1 && idx >= rg.r1 && idx <= rg.r2));
    if (!sel.length) return [idx];
    const out: number[] = [];
    for (const rg of doc.sel.ranges) for (let i = axis === 'col' ? rg.c1 : rg.r1; i <= (axis === 'col' ? rg.c2 : rg.r2) && out.length < 2000; i++) out.push(i);
    return out;
  };

  const onContextMenu = (e: React.MouseEvent) => {
    e.preventDefault();
    const { x, y } = local(e);
    const t = hit(x, y);
    const target: GridTarget =
      t.zone === 'colHeader' || t.zone === 'colResize'
        ? { kind: 'col', c: (t as { c: number }).c }
        : t.zone === 'rowHeader' || t.zone === 'rowResize'
          ? { kind: 'row', r: (t as { r: number }).r }
          : t.zone === 'corner'
            ? { kind: 'corner' }
            : { kind: 'cell', r: (t as { r: number }).r ?? doc.sel.active.r, c: (t as { c: number }).c ?? doc.sel.active.c };
    h.onContextMenu(e.clientX, e.clientY, target);
  };

  const onWheel = (e: React.WheelEvent) => {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      ui.setZoom(zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }
  };

  useEffect(() => {
    // React's onWheel is passive; zoom needs preventDefault
    const sc = scrollerRef.current;
    if (!sc) return;
    const fn = (e: WheelEvent) => {
      if (e.ctrlKey || e.metaKey) e.preventDefault();
    };
    sc.addEventListener('wheel', fn, { passive: false });
    return () => sc.removeEventListener('wheel', fn);
  }, []);

  /* ----------------------------------------------------- keyboard */
  const onSinkKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (composing.current || e.nativeEvent.isComposing) return;
    if (h.onShortcut(e)) {
      e.preventDefault();
      return;
    }
    const mod = isMod(e);
    const shift = e.shiftKey;
    const k = e.key;
    const page = Math.max(1, Math.floor((size.h - headerH) / (sheet.defaultRowHeight * zoom)) - 1);
    const go = (dr: number, dc: number) => {
      e.preventDefault();
      ui.move(dr, dc, shift, mod);
    };
    switch (k) {
      case 'ArrowUp':
        return go(-1, 0);
      case 'ArrowDown':
        return go(1, 0);
      case 'ArrowLeft':
        return go(0, -1);
      case 'ArrowRight':
        return go(0, 1);
      case 'PageDown':
        e.preventDefault();
        if (e.altKey) ui.move(0, 10, shift);
        else for (let i = 0; i < page; i++) ui.move(1, 0, shift);
        return;
      case 'PageUp':
        e.preventDefault();
        if (e.altKey) ui.move(0, -10, shift);
        else for (let i = 0; i < page; i++) ui.move(-1, 0, shift);
        return;
      case 'Home': {
        e.preventDefault();
        const r = mod ? 0 : doc.sel.active.r;
        if (shift) ui.selectCell(r, 0, true);
        else ui.selectCell(Math.max(r, 0), 0);
        return;
      }
      case 'End': {
        if (!mod) return;
        e.preventDefault();
        const x2 = doc.engine.extent(doc.sheet.id);
        ui.selectCell(Math.max(0, x2.rows), Math.max(0, x2.cols), shift);
        return;
      }
      case 'Enter':
        e.preventDefault();
        if (mod || e.altKey) return;
        ui.cycle(shift ? -1 : 1, 0);
        return;
      case 'Tab':
        e.preventDefault();
        ui.cycle(0, shift ? -1 : 1);
        return;
      case 'F2':
        e.preventDefault();
        ui.startEdit(null, 'edit');
        return;
      case 'Delete':
        e.preventDefault();
        doc.clear('contents');
        h.onCommitted?.();
        return;
      case 'Backspace':
        e.preventDefault();
        ui.startEdit('', 'enter');
        return;
      case 'Escape':
        if (doc.clip) {
          doc.clip = null;
          doc.emit('view');
        }
        return;
    }
  };

  const onSinkInput = (e: React.FormEvent<HTMLTextAreaElement>) => {
    if (composing.current) return;
    const s = e.currentTarget;
    const text = s.value.replace(/^ /, '');
    if (!text) return;
    s.value = ' ';
    if (ui.startEdit(text, 'enter')) requestAnimationFrame(() => editorRef.current?.focus());
  };

  // native clipboard events on the focus sink
  useEffect(() => {
    const s = sinkRef.current;
    if (!s) return;
    const copy = (e: ClipboardEvent) => h.onCopy(e, false);
    const cut = (e: ClipboardEvent) => h.onCopy(e, true);
    const paste = (e: ClipboardEvent) => h.onPaste(e);
    s.addEventListener('copy', copy);
    s.addEventListener('cut', cut);
    s.addEventListener('paste', paste);
    return () => {
      s.removeEventListener('copy', copy);
      s.removeEventListener('cut', cut);
      s.removeEventListener('paste', paste);
    };
  }, [h]);

  /* ---------------------------------------------------- cell editor */
  const m = mapperRef.current ?? new Mapper(geo, viewport());
  const editing = edit && edit.sheet === doc.sheet.id ? edit : null;
  const editorBox = useMemo(() => {
    if (!editing || !size.w) return null;
    const rg = ui.withMerges({ r1: editing.r, c1: editing.c, r2: editing.r, c2: editing.c });
    const rect = m.rect(rg);
    const st = doc.styleOf(doc.sheet, editing.r, editing.c);
    const { font, px } = cellFont(st, zoom);
    return { rect, font, px, st };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editing?.r, editing?.c, editing?.sheet, size.w, size.h, zoom, docVersion, uiVersion]);

  const names = doc.wb.names.map((n) => n.name);
  const assist = editing ? analyzeFormula(editing.text, editing.caret, names) : null;
  useEffect(() => setAssistIndex(0), [assist?.prefix?.text]);

  const onEditorKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const el = e.currentTarget;
    const ed = ui.edit;
    if (!ed || e.nativeEvent.isComposing) return;
    const k = e.key;
    const sugg = assist?.suggestions ?? [];
    if (sugg.length) {
      if (k === 'ArrowDown') {
        e.preventDefault();
        setAssistIndex((i) => (i + 1) % sugg.length);
        return;
      }
      if (k === 'ArrowUp') {
        e.preventDefault();
        setAssistIndex((i) => (i - 1 + sugg.length) % sugg.length);
        return;
      }
      if (k === 'Tab' || (k === 'Enter' && !e.altKey && ed.text.startsWith('='))) {
        e.preventDefault();
        const res = applySuggestion(ed.text, assist!, sugg[assistIndex] ?? sugg[0]);
        ui.updateEdit(res.text, res.caret, 'cell');
        requestAnimationFrame(() => el.setSelectionRange(res.caret, res.caret));
        return;
      }
    }
    if (k === 'Enter' && e.altKey) {
      e.preventDefault();
      const s = el.selectionStart;
      const text = ed.text.slice(0, s) + '\n' + ed.text.slice(el.selectionEnd);
      ui.updateEdit(text, s + 1, 'cell');
      requestAnimationFrame(() => el.setSelectionRange(s + 1, s + 1));
      return;
    }
    if (k === 'Enter' || k === 'Tab') {
      e.preventDefault();
      if (!ui.commitEdit({ all: k === 'Enter' && isMod(e) })) return;
      h.onCommitted?.();
      if (!(k === 'Enter' && isMod(e))) {
        if (k === 'Enter') ui.cycle(e.shiftKey ? -1 : 1, 0);
        else ui.cycle(0, e.shiftKey ? -1 : 1);
      }
      focusSink();
      return;
    }
    if (k === 'Escape') {
      e.preventDefault();
      ui.cancelEdit();
      focusSink();
      return;
    }
    if (k === 'F2') {
      e.preventDefault();
      ui.edit = { ...ed, mode: ed.mode === 'enter' ? 'edit' : 'enter' };
      ui.emit();
      return;
    }
    if (k === 'F4') {
      e.preventDefault();
      ui.toggleAbsolute();
      requestAnimationFrame(() => {
        const c = ui.edit?.caret ?? 0;
        el.setSelectionRange(c, c);
      });
      return;
    }
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k] && ed.mode === 'enter') {
      if (ui.canPoint()) {
        e.preventDefault();
        ui.pointMove(arrows[k][0], arrows[k][1], e.shiftKey);
        requestAnimationFrame(() => {
          const c = ui.edit?.caret ?? 0;
          el.setSelectionRange(c, c);
        });
        return;
      }
      e.preventDefault();
      if (!ui.commitEdit()) return;
      h.onCommitted?.();
      ui.move(arrows[k][0], arrows[k][1]);
      focusSink();
    }
  };

  // a freshly opened editor gets focus with the caret where the edit state says (end of text)
  const editKey = editing && editing.source === 'cell' ? `${editing.sheet}:${editing.r}:${editing.c}` : '';
  useLayoutEffect(() => {
    const el = editorRef.current;
    const ed = ui.edit;
    if (!el || !ed || !editKey) return;
    el.value = ed.text;
    el.focus({ preventScroll: true });
    el.setSelectionRange(ed.caret, ed.caret);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editKey]);

  // keep the textarea caret in sync with the model after programmatic changes
  useLayoutEffect(() => {
    const el = editorRef.current;
    const ed = ui.edit;
    if (!el || !ed) return;
    if (el.value !== ed.text) {
      el.value = ed.text;
      el.setSelectionRange(ed.caret, ed.caret);
    }
  });

  const editorStyle = useMemo(() => {
    if (!editorBox || !editing) return null;
    const { rect, font, px, st } = editorBox;
    const lines = editing.text.split('\n').length;
    const minW = rect.w + 2;
    const textW = Math.max(...editing.text.split('\n').map((l) => l.length)) * px * 0.55 + 16;
    const maxW = Math.max(minW, size.w - rect.x - 4);
    const width = Math.min(maxW, Math.max(minW, textW));
    const height = Math.max(rect.h + 1, lines * px * 1.25 + 8);
    return {
      left: rect.x - 1,
      top: rect.y - 1,
      width,
      height,
      font,
      color: st.color ?? undefined,
      background: st.fill ?? undefined,
      textAlign: (st.hAlign === 'center' || st.hAlign === 'right' ? st.hAlign : 'left') as React.CSSProperties['textAlign'],
      whiteSpace: (st.wrap ? 'pre-wrap' : 'pre') as React.CSSProperties['whiteSpace'],
      lineHeight: `${px * 1.22}px`,
    };
  }, [editorBox, editing, size.w]);

  const sinkPos = useMemo(() => {
    const a = doc.sel.active;
    const r = m.rect({ r1: a.r, c1: a.c, r2: a.r, c2: a.c });
    return { left: Math.max(0, Math.min(size.w - 10, r.x)), top: Math.max(0, Math.min(size.h - 10, r.y)) };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [uiVersion, docVersion, size.w, size.h]);

  const editorAnchor = editorStyle && rootRef.current ? (() => {
    const rr = rootRef.current!.getBoundingClientRect();
    return new DOMRect(rr.left + editorStyle.left, rr.top + editorStyle.top, editorStyle.width, editorStyle.height);
  })() : null;

  return (
    <div className={`sheet-grid cursor-${cursor}`} ref={rootRef}>
      <canvas ref={canvasRef} className="grid-canvas" style={{ width: size.w, height: size.h }} />
      <div
        ref={scrollerRef}
        className="grid-scroller"
        onScroll={onScroll}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
        onPointerLeave={() => setHoverNote(null)}
        onDoubleClick={onDoubleClick}
        onContextMenu={onContextMenu}
        onWheel={onWheel}
      >
        <div ref={spacerRef} className="grid-spacer" />
      </div>
      <ObjectLayer doc={doc} ui={ui} mapper={m} width={size.w} height={size.h} headerW={headerW} headerH={headerH} />
      <textarea
        ref={sinkRef}
        className="grid-sink"
        aria-label="Spreadsheet grid"
        style={sinkPos}
        defaultValue=" "
        spellCheck={false}
        autoComplete="off"
        onKeyDown={onSinkKeyDown}
        onInput={onSinkInput}
        onCompositionStart={() => {
          composing.current = true;
        }}
        onCompositionEnd={(e) => {
          composing.current = false;
          const text = e.currentTarget.value.replace(/^ /, '');
          e.currentTarget.value = ' ';
          if (text && ui.startEdit(text, 'enter')) requestAnimationFrame(() => editorRef.current?.focus());
        }}
      />
      {editing && editing.source === 'cell' && editorStyle && (
        <textarea
          ref={editorRef}
          className="cell-editor"
          style={editorStyle}
          defaultValue={editing.text}
          spellCheck={false}
          onChange={(e) => ui.updateEdit(e.target.value, e.target.selectionStart, 'cell')}
          onSelect={(e) => {
            const t = e.currentTarget;
            if (ui.edit && t.selectionStart !== ui.edit.caret) {
              ui.edit = { ...ui.edit, caret: t.selectionStart };
              ui.emit();
            }
          }}
          onKeyDown={onEditorKeyDown}
          onBlur={(e) => {
            const to = e.relatedTarget as HTMLElement | null;
            if (to?.closest('.formula-bar, .formula-assist, .grid-scroller')) return;
            if (ui.edit && !to?.closest('.dialog, .popover, .menu')) {
              if (ui.commitEdit()) h.onCommitted?.();
            }
          }}
          autoFocus
        />
      )}
      {editing && editing.source === 'bar' && editorStyle && (
        <div className="cell-editor ghost" style={editorStyle}>
          {editing.text}
        </div>
      )}
      {editing && editing.source === 'cell' && assist && (
        <FormulaAssist
          info={assist}
          index={assistIndex}
          anchor={editorAnchor}
          onHover={setAssistIndex}
          onPick={(s) => {
            const ed = ui.edit;
            if (!ed || !assist) return;
            const res = applySuggestion(ed.text, assist, s);
            ui.updateEdit(res.text, res.caret, 'cell');
            requestAnimationFrame(() => {
              editorRef.current?.focus();
              editorRef.current?.setSelectionRange(res.caret, res.caret);
            });
          }}
        />
      )}
      {hoverNote && (
        <div className="cell-note-pop" style={{ left: Math.min(hoverNote.x, size.w - 240), top: hoverNote.y }}>
          {hoverNote.text}
        </div>
      )}
    </div>
  );
}
