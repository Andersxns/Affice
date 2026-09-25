import { ChevronLeft, ChevronRight, Eraser, Grid2x2, Highlighter, MonitorUp, MousePointer2, PenLine, Presentation as PresIcon, Square, Timer, X, Zap } from 'lucide-react';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { createPortal } from 'react-dom';
import { api, isElectron } from '@/lib/platform';
import { openContextMenu } from '@/ui/menu';
import type { Presentation, Slide } from '../model';
import { SlideView } from '../render/SlideView';
import { Thumb } from '../editor/panels';
import { animKeyframes, buildOf, hiddenAt, hiddenCss, stepDuration, transitionKeyframes, type Build } from './engine';

/* ================================================================ state */

export interface Stroke {
  tool: 'pen' | 'highlighter';
  color: string;
  width: number;
  pts: [number, number][];
}

export type Tool = 'none' | 'laser' | 'pen' | 'highlighter' | 'eraser';

export interface ShowState {
  index: number;
  done: number;
  playing: number | null;
  seq: number;
  action: { kind: 'enter'; from: number | null; transition: boolean } | { kind: 'step'; step: number } | null;
  screen: 'black' | 'white' | null;
  tool: Tool;
  ink: Record<string, Stroke[]>;
  laser: { x: number; y: number } | null;
  ended: boolean;
  /** The slide's automatic (non-click) animations have played. */
  autoDone: boolean;
}

function useShow(pres: Presentation, start: number, onExit: (i: number) => void, loop = false) {
  const builds = useMemo(() => pres.slides.map((s) => buildOf(s)), [pres]);
  const [st, setSt] = useState<ShowState>(() => ({ index: start, done: 0, playing: null, seq: 1, action: { kind: 'enter', from: null, transition: false }, screen: null, tool: 'none', ink: {}, laser: null, ended: false, autoDone: !builds[start]?.auto }));
  const stRef = useRef(st);
  stRef.current = st;
  const timers = useRef<number[]>([]);
  const later = (fn: () => void, ms: number) => timers.current.push(window.setTimeout(fn, ms));
  useEffect(() => () => timers.current.forEach(clearTimeout), []);

  const visible = useMemo(() => pres.slides.map((s, i) => !s.hidden || i === start), [pres, start]);

  const enter = useCallback(
    (to: number, transition: boolean) => {
      const from = stRef.current.index;
      const b = builds[to];
      setSt((s) => ({ ...s, index: to, done: 0, playing: null, seq: s.seq + 1, action: { kind: 'enter', from, transition }, ended: false, autoDone: !b?.auto }));
      if (b?.auto) {
        // automatic animations run once the slide (and its transition) has appeared
        const tDur = transition ? (pres.slides[to].transition?.dur ?? 700) : 0;
        later(() => {
          setSt((s) => (s.index === to ? { ...s, playing: -1, seq: s.seq + 1, action: { kind: 'step', step: -1 } } : s));
          later(() => setSt((s) => (s.index === to && s.playing === -1 ? { ...s, playing: null, autoDone: true } : s)), stepDuration(b.auto!) + 20);
        }, tDur);
      }
    },
    [builds, pres],
  );

  const nextIndex = (from: number, dir: 1 | -1) => {
    let i = from + dir;
    while (i >= 0 && i < pres.slides.length && !visible[i]) i += dir;
    return i;
  };

  const next = useCallback(() => {
    const s = stRef.current;
    if (s.screen) return setSt((x) => ({ ...x, screen: null }));
    if (s.ended) return onExit(s.index);
    const b = builds[s.index];
    if (b && s.done < b.steps.length) {
      const step = s.done;
      setSt((x) => ({ ...x, done: x.done + 1, playing: step, seq: x.seq + 1, action: { kind: 'step', step } }));
      later(() => setSt((x) => (x.playing === step && x.index === s.index ? { ...x, playing: null } : x)), stepDuration(b.steps[step]) + 20);
      return;
    }
    const n = nextIndex(s.index, 1);
    if (n >= pres.slides.length) {
      if (loop) enter(nextIndex(-1, 1), true);
      else setSt((x) => ({ ...x, ended: true, seq: x.seq + 1 }));
      return;
    }
    enter(n, true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [builds, pres, enter, onExit, loop]);

  const prev = useCallback(() => {
    const s = stRef.current;
    if (s.screen) return setSt((x) => ({ ...x, screen: null }));
    if (s.ended) return setSt((x) => ({ ...x, ended: false, seq: x.seq + 1 }));
    if (s.done > 0) return setSt((x) => ({ ...x, done: x.done - 1, playing: null, seq: x.seq + 1, action: null }));
    const p = nextIndex(s.index, -1);
    if (p < 0) return;
    setSt((x) => ({ ...x, index: p, done: builds[p]?.steps.length ?? 0, playing: null, seq: x.seq + 1, action: { kind: 'enter', from: s.index, transition: false }, autoDone: true }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [builds]);

  const goto = useCallback((i: number) => enter(Math.max(0, Math.min(pres.slides.length - 1, i)), false), [enter, pres]);

  // automatic advance ("After" timing)
  useEffect(() => {
    const t = pres.slides[st.index]?.transition;
    if (!t?.after || st.ended) return;
    const id = window.setTimeout(next, t.after);
    return () => clearTimeout(id);
  }, [st.index, st.done, st.ended, pres, next]);

  return { st, setSt, builds, next, prev, goto };
}

type ShowApi = ReturnType<typeof useShow>;

/* ================================================================ stage */

function useSize(ref: React.RefObject<HTMLElement | null>): { w: number; h: number } {
  const [size, setSize] = useState({ w: 0, h: 0 });
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, [ref]);
  return size;
}

let stageSeq = 0;

interface StageProps {
  pres: Presentation;
  st: ShowState;
  build: Build;
  /** Pointer interaction (clicks advance, pen draws). */
  interactive?: boolean;
  onClick?: () => void;
  onInk?: (slideId: string, strokes: Stroke[]) => void;
  onLaser?: (p: { x: number; y: number } | null) => void;
  onLink?: (href: string) => void;
  /** Show the slide at a given index, frozen after `done` click steps (the presenter's "next" preview). */
  still?: { index: number; done?: number };
  className?: string;
}

/** Renders one slide scaled to fit its box and plays transitions and animations. */
export function ShowStage({ pres, st, build, interactive, onClick, onInk, onLaser, onLink, still, className = '' }: StageProps) {
  const boxRef = useRef<HTMLDivElement>(null);
  const layerIn = useRef<HTMLDivElement>(null);
  const layerOut = useRef<HTMLDivElement>(null);
  const size = useSize(boxRef);
  const scope = useMemo(() => `show-scope-${++stageSeq}`, []);
  const [outgoing, setOutgoing] = useState<number | null>(null);
  const drawing = useRef<Stroke | null>(null);
  const [live, setLive] = useState<Stroke | null>(null);
  const index = still ? still.index : st.index;
  const slide: Slide | undefined = pres.slides[index];
  const scale = size.w && size.h ? Math.min(size.w / pres.size.w, size.h / pres.size.h) : 0;
  const done = still ? (still.done ?? 0) : st.done;
  const hidden = useMemo(() => {
    if (!slide) return new Set<string>();
    if (still) return hiddenAt(buildOf(slide), done, null, true);
    return hiddenAt(build, done, st.playing, st.autoDone || st.playing === -1);
  }, [slide, build, done, st.playing, st.autoDone, still]);

  // transitions and animations run on each action
  useLayoutEffect(() => {
    if (still || !st.action || !scale) return;
    const a = st.action;
    if (a.kind === 'enter') {
      const t = slide?.transition;
      if (a.transition && a.from !== null && a.from !== index && t && t.type !== 'none') {
        setOutgoing(a.from);
      } else setOutgoing(null);
    } else if (a.kind === 'step') {
      const items = a.step === -1 ? build.auto : build.steps[a.step];
      if (!items || !layerIn.current) return;
      const anims: Animation[] = [];
      for (const it of items) {
        const host = layerIn.current.querySelector<HTMLElement>(`[data-el="${CSS.escape(it.anim.el)}"]`);
        if (!host) continue;
        const target = it.para !== undefined ? host.querySelectorAll<HTMLElement>('.sl-p')[it.para] : host;
        if (!target) continue;
        const el = slide?.elements.find((e) => e.id === it.anim.el);
        const base = it.para !== undefined ? '' : host.style.transform;
        const k = animKeyframes(it.anim, base, el ?? { x: 0, y: 0, w: 0, h: 0 }, pres);
        anims.push(target.animate(k.frames, { duration: k.duration, delay: it.offset, easing: k.easing, fill: it.anim.cls === 'emph' ? 'none' : 'both', iterations: Math.max(1, it.anim.repeat ?? 1) }));
      }
      return () => {
        for (const x of anims) if (x.playState === 'running') x.finish();
      };
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [st.seq, scale]);

  // run the slide transition once both layers are in the DOM
  useLayoutEffect(() => {
    if (outgoing === null || !slide || !layerIn.current) return;
    const t = slide.transition!;
    const k = transitionKeyframes(t);
    const dur = t.dur ?? 700;
    const anims: Animation[] = [];
    if (k.in) anims.push(layerIn.current.animate(k.in, { duration: dur, easing: k.easing, fill: 'both' }));
    if (k.out && layerOut.current) anims.push(layerOut.current.animate(k.out, { duration: dur, easing: k.easing, fill: 'both' }));
    if (t.type === 'morph' && layerOut.current) {
      const oldSlide = pres.slides[outgoing];
      const byName = new Map(oldSlide.elements.filter((e) => e.name).map((e) => [`${e.type}:${e.name}`, e]));
      for (const e of slide.elements) {
        const node = layerIn.current.querySelector<HTMLElement>(`[data-el="${CSS.escape(e.id)}"]`);
        if (!node) continue;
        const o = byName.get(`${e.type}:${e.name}`);
        if (!o) {
          anims.push(node.animate([{ opacity: 0 }, { opacity: 1 }], { duration: dur, easing: 'ease', fill: 'both' }));
          continue;
        }
        const oldNode = layerOut.current.querySelector<HTMLElement>(`[data-el="${CSS.escape(o.id)}"]`);
        if (oldNode) oldNode.style.visibility = 'hidden';
        const dx = o.x + o.w / 2 - (e.x + e.w / 2);
        const dy = o.y + o.h / 2 - (e.y + e.h / 2);
        const sx = e.w ? o.w / e.w : 1;
        const sy = e.h ? o.h / e.h : 1;
        anims.push(
          node.animate(
            [
              { transform: `translate(${dx}px, ${dy}px) rotate(${o.rot ?? 0}deg) scale(${sx}, ${sy})`, opacity: o.opacity ?? 1 },
              { transform: `translate(0px, 0px) rotate(${e.rot ?? 0}deg) scale(1, 1)`, opacity: e.opacity ?? 1 },
            ],
            { duration: dur, easing: 'cubic-bezier(.4,0,.2,1)', fill: 'both' },
          ),
        );
      }
    }
    const id = window.setTimeout(() => setOutgoing(null), dur + 30);
    return () => {
      clearTimeout(id);
      anims.forEach((x) => x.cancel());
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [outgoing]);

  const toSlide = (e: React.PointerEvent): [number, number] => {
    const r = layerIn.current!.getBoundingClientRect();
    return [(e.clientX - r.left) / scale, (e.clientY - r.top) / scale];
  };

  const ink = slide ? (st.ink[slide.id] ?? []) : [];
  const drawTool = interactive && (st.tool === 'pen' || st.tool === 'highlighter');

  const onPointerDown = (e: React.PointerEvent) => {
    if (!interactive || e.button !== 0 || !slide) return;
    if (drawTool) {
      const [x, y] = toSlide(e);
      const s: Stroke = st.tool === 'highlighter' ? { tool: 'highlighter', color: '#ffe600', width: 22, pts: [[x, y]] } : { tool: 'pen', color: '#e5484d', width: 4, pts: [[x, y]] };
      drawing.current = s;
      setLive(s);
      (e.target as HTMLElement).setPointerCapture?.(e.pointerId);
      return;
    }
    if (st.tool === 'eraser') {
      const hit = (e.target as Element).closest('[data-stroke]');
      if (hit) onInk?.(slide.id, ink.filter((_, i) => i !== Number(hit.getAttribute('data-stroke'))));
      return;
    }
    const linkEl = (e.target as HTMLElement).closest<HTMLElement>('[data-el]');
    const id = linkEl?.getAttribute('data-el');
    const target = id ? slide.elements.find((x) => x.id === id) : undefined;
    if (target?.link) {
      onLink?.(target.link);
      return;
    }
    const a = (e.target as HTMLElement).closest('a.sl-link') as HTMLAnchorElement | null;
    if (a?.getAttribute('href')) {
      onLink?.(a.getAttribute('href')!);
      return;
    }
    onClick?.();
  };
  const onPointerMove = (e: React.PointerEvent) => {
    if (!interactive || !scale) return;
    if (st.tool === 'laser') {
      const [x, y] = toSlide(e);
      onLaser?.({ x, y });
    }
    const d = drawing.current;
    if (d) {
      const [x, y] = toSlide(e);
      d.pts.push([x, y]);
      setLive({ ...d });
    }
  };
  const onPointerUp = () => {
    const d = drawing.current;
    drawing.current = null;
    setLive(null);
    if (d && slide && d.pts.length > 1) onInk?.(slide.id, [...ink, d]);
  };

  const strokePath = (s: Stroke) => s.pts.map((p, i) => `${i ? 'L' : 'M'}${Math.round(p[0] * 10) / 10} ${Math.round(p[1] * 10) / 10}`).join('');
  const layerStyle: CSSProperties = { width: pres.size.w, height: pres.size.h, transform: `scale(${scale})` };
  const stageStyle: CSSProperties = { width: pres.size.w * scale, height: pres.size.h * scale };
  const outSlide = outgoing !== null ? pres.slides[outgoing] : null;
  const outOnTop = slide?.transition ? transitionKeyframes(slide.transition).outOnTop : false;

  return (
    <div ref={boxRef} className={`show-box ${className}${st.tool === 'laser' && interactive ? ' laser-on' : ''}${drawTool ? ' pen-on' : ''}${st.tool === 'eraser' && interactive ? ' eraser-on' : ''}`} onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerLeave={() => interactive && st.tool === 'laser' && onLaser?.(null)}>
      {scale > 0 && slide && (
        <div className={`show-stage ${scope}`} style={stageStyle}>
          <style>{hiddenCss(scope, hidden)}</style>
          {outSlide && (
            <div className={`show-layer out${outOnTop ? ' on-top' : ''}`} ref={layerOut} style={layerStyle}>
              <SlideView pres={pres} slide={outSlide} slideNumber={outgoing! + 1} mode="show" />
            </div>
          )}
          <div className="show-layer in" ref={layerIn} style={layerStyle} key={slide.id}>
            <SlideView pres={pres} slide={slide} slideNumber={index + 1} mode="show" />
            {(ink.length > 0 || live) && (
              <svg className="show-ink" width={pres.size.w} height={pres.size.h} viewBox={`0 0 ${pres.size.w} ${pres.size.h}`}>
                {[...ink, ...(live ? [live] : [])].map((s, i) => (
                  <path key={i} data-stroke={i} d={strokePath(s)} fill="none" stroke={s.color} strokeWidth={s.width} strokeLinecap="round" strokeLinejoin="round" opacity={s.tool === 'highlighter' ? 0.42 : 1} />
                ))}
              </svg>
            )}
            {st.laser && !still && <div className="show-laser" style={{ left: st.laser.x, top: st.laser.y, transform: `translate(-50%, -50%) scale(${1 / Math.max(scale, 0.2)})` }} />}
          </div>
          {st.screen && !still && <div className={`show-screen ${st.screen}`} />}
          {st.ended && !still && (
            <div className="show-screen black end">
              <span>End of slide show. Click to exit.</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ============================================================ controls */

function ToolButtons({ show }: { show: ShowApi }) {
  const { st, setSt } = show;
  const tool = (t: Tool) => setSt((s) => ({ ...s, tool: s.tool === t ? 'none' : t, laser: null }));
  return (
    <>
      <button className={`show-btn${st.tool === 'laser' ? ' on' : ''}`} data-tip="Laser pointer (Ctrl+L)" onClick={() => tool('laser')}>
        <Zap size={16} />
      </button>
      <button className={`show-btn${st.tool === 'pen' ? ' on' : ''}`} data-tip="Pen (Ctrl+P)" onClick={() => tool('pen')}>
        <PenLine size={16} />
      </button>
      <button className={`show-btn${st.tool === 'highlighter' ? ' on' : ''}`} data-tip="Highlighter (Ctrl+I)" onClick={() => tool('highlighter')}>
        <Highlighter size={16} />
      </button>
      <button className={`show-btn${st.tool === 'eraser' ? ' on' : ''}`} data-tip="Eraser (Ctrl+E) · press E to erase all" onClick={() => tool('eraser')}>
        <Eraser size={16} />
      </button>
      <button className={`show-btn${st.tool === 'none' ? ' on' : ''}`} data-tip="Arrow (Ctrl+A)" onClick={() => setSt((s) => ({ ...s, tool: 'none', laser: null }))}>
        <MousePointer2 size={16} />
      </button>
    </>
  );
}

function SlideGrid({ pres, current, onPick, onClose }: { pres: Presentation; current: number; onPick: (i: number) => void; onClose: () => void }) {
  return (
    <div className="show-grid thin-scroll" onClick={onClose}>
      <div className="show-grid-inner">
        {pres.slides.map((s, i) => (
          <button
            key={s.id}
            type="button"
            className={`show-grid-item${i === current ? ' current' : ''}${s.hidden ? ' hidden-slide' : ''}`}
            onClick={(e) => {
              e.stopPropagation();
              onPick(i);
            }}
          >
            <Thumb pres={pres} slide={s} index={i} width={220} />
            <span>{i + 1}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

function useKeys(show: ShowApi, pres: Presentation, onExit: () => void, extra: { grid: () => void }, target: Window = window) {
  const typed = useRef('');
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const k = e.key;
      const mod = e.ctrlKey || e.metaKey;
      const { next, prev, goto, setSt, st } = show;
      const handled = () => {
        e.preventDefault();
        e.stopPropagation();
      };
      if (/^\d$/.test(k) && !mod) {
        typed.current += k;
        return handled();
      }
      if (k === 'Enter' && typed.current) {
        goto(parseInt(typed.current, 10) - 1);
        typed.current = '';
        return handled();
      }
      typed.current = '';
      if (mod && ['p', 'l', 'i', 'e', 'a'].includes(k.toLowerCase())) {
        const t: Record<string, Tool> = { p: 'pen', l: 'laser', i: 'highlighter', e: 'eraser', a: 'none' };
        setSt((s) => ({ ...s, tool: t[k.toLowerCase()], laser: null }));
        return handled();
      }
      switch (k) {
        case 'ArrowRight':
        case 'ArrowDown':
        case 'PageDown':
        case ' ':
        case 'n':
        case 'N':
        case 'Enter':
          handled();
          return next();
        case 'ArrowLeft':
        case 'ArrowUp':
        case 'PageUp':
        case 'Backspace':
        case 'p':
        case 'P':
          handled();
          return prev();
        case 'Home':
          handled();
          return goto(0);
        case 'End':
          handled();
          return goto(pres.slides.length - 1);
        case 'b':
        case 'B':
        case '.':
          handled();
          return setSt((s) => ({ ...s, screen: s.screen === 'black' ? null : 'black' }));
        case 'w':
        case 'W':
        case ',':
          handled();
          return setSt((s) => ({ ...s, screen: s.screen === 'white' ? null : 'white' }));
        case 'e':
        case 'E': {
          handled();
          const id = pres.slides[st.index]?.id;
          if (id) setSt((s) => ({ ...s, ink: { ...s.ink, [id]: [] } }));
          return;
        }
        case 'g':
        case 'G':
        case '-':
          handled();
          return extra.grid();
        case 'Escape':
          handled();
          if (st.tool !== 'none') return setSt((s) => ({ ...s, tool: 'none', laser: null }));
          return onExit();
      }
    };
    target.addEventListener('keydown', onKey, true);
    return () => target.removeEventListener('keydown', onKey, true);
  }, [show, pres, onExit, extra, target]);
}

function useLinks(show: ShowApi, pres: Presentation) {
  return useCallback(
    (href: string) => {
      const { next, prev, goto } = show;
      if (href === '#next') return next();
      if (href === '#prev') return prev();
      if (href === '#first') return goto(0);
      if (href === '#last') return goto(pres.slides.length - 1);
      const m = /^#slide:(\d+)$/.exec(href);
      if (m) return goto(parseInt(m[1], 10) - 1);
      if (/^(https?:|mailto:)/i.test(href)) api.app.openExternal(href);
    },
    [show, pres],
  );
}

/* ============================================================ full show */

export interface SlideshowProps {
  pres: Presentation;
  start: number;
  presenter: boolean;
  onExit: (index: number) => void;
}

export function Slideshow({ pres, start, presenter, onExit }: SlideshowProps) {
  const exit = useRef(onExit);
  exit.current = onExit;
  const show = useShow(pres, start, (i) => exit.current(i));
  const [grid, setGrid] = useState(false);
  const [chrome, setChrome] = useState(false);
  const hideTimer = useRef(0);
  const doExit = useCallback(() => exit.current(show.st.index), [show.st.index]);
  const onLink = useLinks(show, pres);
  const setInk = (id: string, strokes: Stroke[]) => show.setSt((s) => ({ ...s, ink: { ...s.ink, [id]: strokes } }));
  const setLaser = (p: { x: number; y: number } | null) => show.setSt((s) => ({ ...s, laser: p }));

  // full screen for the audience view
  useEffect(() => {
    if (presenter) return;
    let was = false;
    void (async () => {
      if (isElectron) {
        was = await api.window.isFullScreen();
        if (!was) api.window.setFullScreen(true);
      } else void document.documentElement.requestFullscreen?.().catch(() => undefined);
    })();
    return () => {
      if (isElectron) {
        if (!was) api.window.setFullScreen(false);
      } else if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    };
  }, [presenter]);

  useKeys(show, pres, doExit, useMemo(() => ({ grid: () => setGrid((g) => !g) }), []));

  const wake = () => {
    setChrome(true);
    clearTimeout(hideTimer.current);
    hideTimer.current = window.setTimeout(() => setChrome(false), 2600);
  };

  const menu = (e: React.MouseEvent) => {
    e.preventDefault();
    openContextMenu(e, [
      { label: 'Next', shortcut: '→', onSelect: show.next },
      { label: 'Previous', shortcut: '←', onSelect: show.prev },
      { label: 'See all slides', shortcut: 'G', onSelect: () => setGrid(true) },
      { separator: true },
      { label: 'Laser pointer', onSelect: () => show.setSt((s) => ({ ...s, tool: 'laser' })) },
      { label: 'Pen', onSelect: () => show.setSt((s) => ({ ...s, tool: 'pen' })) },
      { label: 'Highlighter', onSelect: () => show.setSt((s) => ({ ...s, tool: 'highlighter' })) },
      { label: 'Erase all ink on slide', shortcut: 'E', onSelect: () => setInk(pres.slides[show.st.index].id, []) },
      { separator: true },
      { label: 'Black screen', shortcut: 'B', onSelect: () => show.setSt((s) => ({ ...s, screen: 'black' })) },
      { label: 'White screen', shortcut: 'W', onSelect: () => show.setSt((s) => ({ ...s, screen: 'white' })) },
      { separator: true },
      { label: 'End show', shortcut: 'Esc', onSelect: doExit },
    ]);
  };

  if (presenter) return <PresenterView pres={pres} show={show} onExit={doExit} onLink={onLink} setInk={setInk} setLaser={setLaser} />;

  return createPortal(
    <div className={`show-root${chrome ? ' chrome' : ''}`} onMouseMove={wake} onContextMenu={menu}>
      <ShowStage pres={pres} st={show.st} build={show.builds[show.st.index]} interactive onClick={show.next} onInk={setInk} onLaser={setLaser} onLink={onLink} />
      <div className="show-toolbar" onPointerDown={(e) => e.stopPropagation()}>
        <button className="show-btn" data-tip="Previous" onClick={show.prev}>
          <ChevronLeft size={18} />
        </button>
        <button className="show-btn" data-tip="Next" onClick={show.next}>
          <ChevronRight size={18} />
        </button>
        <ToolButtons show={show} />
        <button className="show-btn" data-tip="See all slides (G)" onClick={() => setGrid(true)}>
          <Grid2x2 size={16} />
        </button>
        <button className="show-btn" data-tip="End show (Esc)" onClick={doExit}>
          <X size={16} />
        </button>
        <span className="show-count">
          {show.st.index + 1} / {pres.slides.length}
        </span>
      </div>
      {grid && <SlideGrid pres={pres} current={show.st.index} onPick={(i) => (show.goto(i), setGrid(false))} onClose={() => setGrid(false)} />}
    </div>,
    document.body,
  );
}

/* ======================================================= presenter view */

function useAudienceWindow(open: boolean): HTMLElement | null {
  const [host, setHost] = useState<HTMLElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const w = window.open('', 'affice-audience', 'popup,width=1280,height=720');
    if (!w) return;
    w.document.title = 'Affice — Slide show';
    w.document.body.innerHTML = '';
    for (const n of Array.from(document.querySelectorAll('style, link[rel="stylesheet"]'))) w.document.head.appendChild(n.cloneNode(true));
    w.document.documentElement.dataset.theme = 'dark';
    w.document.body.className = 'audience-body';
    const div = w.document.createElement('div');
    div.className = 'audience-root';
    w.document.body.appendChild(div);
    setHost(div);
    const onClose = () => setHost(null);
    w.addEventListener('beforeunload', onClose);
    return () => {
      w.removeEventListener('beforeunload', onClose);
      w.close();
      setHost(null);
    };
  }, [open]);
  return host;
}

function Clock() {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);
  return <span>{now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>;
}

function Stopwatch() {
  const [elapsed, setElapsed] = useState(0);
  const [running, setRunning] = useState(true);
  const last = useRef(Date.now());
  useEffect(() => {
    const id = setInterval(() => {
      const now = Date.now();
      if (running) setElapsed((e) => e + (now - last.current));
      last.current = now;
    }, 500);
    return () => clearInterval(id);
  }, [running]);
  const s = Math.floor(elapsed / 1000);
  const fmt = `${Math.floor(s / 3600)}:${String(Math.floor((s % 3600) / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`;
  return (
    <span className="presenter-timer">
      <Timer size={15} /> {fmt}
      <button className="show-btn small" onClick={() => setRunning((r) => !r)}>
        {running ? 'Pause' : 'Resume'}
      </button>
      <button className="show-btn small" onClick={() => setElapsed(0)}>
        Reset
      </button>
    </span>
  );
}

function PresenterView({ pres, show, onExit, onLink, setInk, setLaser }: { pres: Presentation; show: ShowApi; onExit: () => void; onLink: (h: string) => void; setInk: (id: string, s: Stroke[]) => void; setLaser: (p: { x: number; y: number } | null) => void }) {
  const [audience, setAudience] = useState(isElectron);
  const host = useAudienceWindow(audience);
  const [grid, setGrid] = useState(false);
  const [notesSize, setNotesSize] = useState(20);
  const { st } = show;
  const slide = pres.slides[st.index];
  const nextIdx = (() => {
    const b = show.builds[st.index];
    if (b && st.done < b.steps.length) return st.index;
    let i = st.index + 1;
    while (i < pres.slides.length && pres.slides[i].hidden) i++;
    return i < pres.slides.length ? i : null;
  })();
  useKeys(show, pres, onExit, useMemo(() => ({ grid: () => setGrid((g) => !g) }), []));
  const audienceWin = host?.ownerDocument.defaultView ?? null;
  return createPortal(
    <div className="presenter-root">
      <div className="presenter-top">
        <PresIcon size={16} /> <strong>Presenter view</strong>
        <span className="spacer" />
        <Stopwatch />
        <Clock />
        <button className="show-btn small" onClick={() => setAudience((a) => !a)} data-tip="Show the slides in a separate window you can move to the projector">
          <MonitorUp size={15} /> {host ? 'Close audience window' : 'Open audience window'}
        </button>
        <button className="show-btn small danger" onClick={onExit}>
          End show
        </button>
      </div>
      <div className="presenter-main">
        <div className="presenter-current">
          <ShowStage pres={pres} st={st} build={show.builds[st.index]} interactive onClick={show.next} onInk={setInk} onLaser={setLaser} onLink={onLink} />
          <div className="presenter-bar">
            <button className="show-btn" onClick={show.prev} data-tip="Previous">
              <ChevronLeft size={18} />
            </button>
            <span className="show-count">
              Slide {st.index + 1} of {pres.slides.length}
            </span>
            <button className="show-btn" onClick={show.next} data-tip="Next">
              <ChevronRight size={18} />
            </button>
            <span className="presenter-sep" />
            <ToolButtons show={show} />
            <button className="show-btn" data-tip="See all slides" onClick={() => setGrid(true)}>
              <Grid2x2 size={16} />
            </button>
            <button className={`show-btn${st.screen === 'black' ? ' on' : ''}`} data-tip="Black screen (B)" onClick={() => show.setSt((s) => ({ ...s, screen: s.screen === 'black' ? null : 'black' }))}>
              <Square size={16} fill="currentColor" />
            </button>
          </div>
        </div>
        <div className="presenter-side">
          <div className="presenter-next">
            <div className="presenter-label">{nextIdx === st.index ? 'Next: animation on this slide' : nextIdx === null ? 'End of slide show' : `Next: slide ${nextIdx + 1}`}</div>
            {nextIdx !== null && <ShowStage pres={pres} st={st} build={show.builds[nextIdx]} still={{ index: nextIdx, done: nextIdx === st.index ? st.done + 1 : 0 }} className="small" />}
          </div>
          <div className="presenter-notes">
            <div className="presenter-label">
              Notes
              <span className="spacer" />
              <button className="show-btn small" onClick={() => setNotesSize((s) => Math.max(12, s - 2))}>
                A−
              </button>
              <button className="show-btn small" onClick={() => setNotesSize((s) => Math.min(48, s + 2))}>
                A+
              </button>
            </div>
            <div className="presenter-notes-text thin-scroll" style={{ fontSize: notesSize }}>
              {slide?.notes?.trim() ? slide.notes : <span className="muted">No notes for this slide.</span>}
            </div>
          </div>
        </div>
      </div>
      {grid && <SlideGrid pres={pres} current={st.index} onPick={(i) => (show.goto(i), setGrid(false))} onClose={() => setGrid(false)} />}
      {host &&
        createPortal(
          <AudienceKeys show={show} pres={pres} onExit={onExit} win={audienceWin}>
            <ShowStage pres={pres} st={st} build={show.builds[st.index]} interactive onClick={show.next} onInk={setInk} onLaser={setLaser} onLink={onLink} />
          </AudienceKeys>,
          host,
        )}
    </div>,
    document.body,
  );
}

function AudienceKeys({ show, pres, onExit, win, children }: { show: ShowApi; pres: Presentation; onExit: () => void; win: Window | null; children: React.ReactNode }) {
  useKeys(show, pres, onExit, useMemo(() => ({ grid: () => undefined }), []), win ?? window);
  return <div className="show-root audience">{children}</div>;
}

/* ============================================================== preview */

/** Plays the current slide's transition (from the previous slide) or its animations over the editor. */
export function SlidePreview({ pres, index, what, onDone }: { pres: Presentation; index: number; what: 'transition' | 'anims'; onDone: () => void }) {
  const builds = useMemo(() => pres.slides.map((s) => buildOf(s)), [pres]);
  const b = builds[index];
  const [st, setSt] = useState<ShowState>(() => ({ index: what === 'transition' && index > 0 ? index - 1 : index, done: what === 'transition' && index > 0 ? (builds[index - 1]?.steps.length ?? 0) : 0, playing: null, seq: 1, action: null, screen: null, tool: 'none', ink: {}, laser: null, ended: false, autoDone: what === 'transition' }));
  useEffect(() => {
    const timers: number[] = [];
    const at = (ms: number, fn: () => void) => timers.push(window.setTimeout(fn, ms));
    let t = 250;
    if (what === 'transition') {
      at(t, () => setSt((s) => ({ ...s, index, done: 0, seq: s.seq + 1, action: { kind: 'enter', from: index > 0 ? index - 1 : null, transition: true }, autoDone: true })));
      t += (pres.slides[index].transition?.dur ?? 700) + 400;
    } else {
      if (b.auto) {
        at(t, () => setSt((s) => ({ ...s, playing: -1, seq: s.seq + 1, action: { kind: 'step', step: -1 } })));
        t += stepDuration(b.auto) + 150;
        at(t - 100, () => setSt((s) => ({ ...s, playing: null, autoDone: true })));
      }
      b.steps.forEach((step, i) => {
        at(t, () => setSt((s) => ({ ...s, done: i + 1, playing: i, seq: s.seq + 1, action: { kind: 'step', step: i } })));
        t += stepDuration(step) + 350;
      });
      t += 300;
    }
    at(t, onDone);
    return () => timers.forEach(clearTimeout);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return (
    <div className="slide-preview" onClick={onDone}>
      <ShowStage pres={pres} st={st} build={builds[st.index]} />
    </div>
  );
}
