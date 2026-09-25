import { ArrowDown, ArrowUp, ChevronDown, ChevronUp, MousePointerClick, Play, Replace, Search, Sparkles, Timer, Trash2, X } from 'lucide-react';
import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import { Checkbox } from '@/ui/controls';
import type { SlidesDoc } from '../doc';
import { findEl, plainText, walkEls, type Anim, type El, type Para } from '../model';
import { moveSlides } from '../ops';
import { ANIMATIONS } from './SlidesRibbon';
import { Thumb } from './panels';

/* ======================================================== animation pane */

function elLabel(e: El | undefined): string {
  if (!e) return 'Missing object';
  const txt = e.type === 'shape' ? plainText(e.text).trim().split('\n')[0] : '';
  return `${e.name ?? e.type}${txt ? `: ${txt.slice(0, 28)}${txt.length > 28 ? '…' : ''}` : ''}`;
}

export function AnimationPane({ doc, onClose, onPlay }: { doc: SlidesDoc; onClose: () => void; onPlay: () => void }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  const slide = doc.slide;
  const anims = slide?.anims ?? [];
  let click = 0;
  const move = (i: number, d: -1 | 1) => {
    const j = i + d;
    if (j < 0 || j >= anims.length) return;
    const next = [...anims];
    [next[i], next[j]] = [next[j], next[i]];
    doc.updateSlide('Reorder animation', (s) => ({ ...s, anims: next }));
  };
  const patch = (id: string, p: Partial<Anim>) => doc.updateSlide('Animation', (s) => ({ ...s, anims: (s.anims ?? []).map((a) => (a.id === id ? { ...a, ...p } : a)) }));
  return (
    <aside className="sl-side-pane">
      <div className="sl-side-head">
        <Sparkles size={15} /> Animation pane
        <span className="spacer" />
        <button className="btn btn-sm" onClick={onPlay} disabled={!anims.length}>
          <Play size={13} /> Play all
        </button>
        <button className="icon-btn icon-btn-sm" aria-label="Close" onClick={onClose}>
          <X size={15} />
        </button>
      </div>
      <div className="sl-anim-list thin-scroll">
        {!anims.length && <p className="muted small pad">Select an object and pick an effect on the Animations tab.</p>}
        {anims.map((a, i) => {
          if (a.start === 'click') click++;
          const def = ANIMATIONS.find((x) => x.cls === a.cls && x.effect === a.effect);
          const el = findEl(slide.elements, a.el);
          const selected = doc.sel.els.includes(a.el);
          return (
            <div key={a.id} className={`sl-anim-row${selected ? ' selected' : ''}`} onClick={() => doc.selectEls([a.el])}>
              <span className="sl-anim-step">{a.start === 'click' ? click : ''}</span>
              <span className={`sl-anim-dot anim-${a.cls}`} />
              <div className="grow ellipsis">
                <div className="ellipsis">{elLabel(el)}</div>
                <div className="muted small">
                  {def?.label ?? a.effect}
                  {a.byPara ? ' · by paragraph' : ''}
                </div>
              </div>
              <select className="select select-sm" value={a.start} onClick={(e) => e.stopPropagation()} onChange={(e) => patch(a.id, { start: e.target.value as Anim['start'] })} aria-label="Start">
                <option value="click">On click</option>
                <option value="with">With previous</option>
                <option value="after">After previous</option>
              </select>
              <div className="sl-anim-btns" onClick={(e) => e.stopPropagation()}>
                <button className="icon-btn icon-btn-sm" aria-label="Move up" onClick={() => move(i, -1)} disabled={i === 0}>
                  <ArrowUp size={13} />
                </button>
                <button className="icon-btn icon-btn-sm" aria-label="Move down" onClick={() => move(i, 1)} disabled={i === anims.length - 1}>
                  <ArrowDown size={13} />
                </button>
                <button className="icon-btn icon-btn-sm" aria-label="Remove" onClick={() => doc.updateSlide('Remove animation', (s) => ({ ...s, anims: (s.anims ?? []).filter((x) => x.id !== a.id) }))}>
                  <Trash2 size={13} />
                </button>
              </div>
              {el?.type === 'shape' && (el.text?.paras.length ?? 0) > 1 && (
                <div className="sl-anim-opts" onClick={(e) => e.stopPropagation()}>
                  <Checkbox checked={!!a.byPara} onChange={(v) => patch(a.id, { byPara: v || undefined })} label="By paragraph" />
                </div>
              )}
            </div>
          );
        })}
      </div>
      <div className="sl-side-foot muted small">
        <MousePointerClick size={13} /> On click · <Timer size={13} /> With / After previous
      </div>
    </aside>
  );
}

/* ================================================================ find */

interface Hit {
  slide: number;
  el: string;
  cell?: [number, number];
  para: number;
  start: number;
  len: number;
}

function paraText(p: Para): string {
  return p.runs.map((r) => r.text).join('');
}

function findHits(doc: SlidesDoc, q: string, matchCase: boolean, whole: boolean): Hit[] {
  if (!q) return [];
  const hits: Hit[] = [];
  const needle = matchCase ? q : q.toLowerCase();
  const esc = needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(whole ? `\\b${esc}\\b` : esc, matchCase ? 'g' : 'gi');
  doc.pres.slides.forEach((s, si) => {
    walkEls(s.elements, (e) => {
      const scan = (paras: Para[], cell?: [number, number]) =>
        paras.forEach((p, pi) => {
          const t = paraText(p);
          for (const m of t.matchAll(re)) hits.push({ slide: si, el: e.id, cell, para: pi, start: m.index ?? 0, len: m[0].length });
        });
      if (e.type === 'shape' && e.text) scan(e.text.paras);
      if (e.type === 'table') e.rows.forEach((r, ri) => r.cells.forEach((c, ci) => scan(c.text.paras, [ri, ci])));
    });
  });
  return hits;
}

/** Replaces [start, start+len) in a paragraph, keeping the formatting of the first affected run. */
function replaceInPara(p: Para, start: number, len: number, text: string): Para {
  const runs = p.runs.map((r) => ({ ...r }));
  let pos = 0;
  let inserted = false;
  const out: typeof runs = [];
  for (const r of runs) {
    const a = pos;
    const b = pos + r.text.length;
    pos = b;
    if (b <= start || a >= start + len) {
      out.push(r);
      continue;
    }
    const keepL = r.text.slice(0, Math.max(0, start - a));
    const keepR = r.text.slice(Math.max(0, start + len - a));
    const mid = inserted ? '' : text;
    inserted = true;
    const t = keepL + mid + keepR;
    if (t) out.push({ ...r, text: t });
  }
  return { ...p, runs: out };
}

export function SlideFind({ doc, mode, onClose }: { doc: SlidesDoc; mode: 'find' | 'replace'; onClose: () => void }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  const [q, setQ] = useState('');
  const [rep, setRep] = useState('');
  const [showRep, setShowRep] = useState(mode === 'replace');
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [cur, setCur] = useState(0);
  const input = useRef<HTMLInputElement>(null);
  const hits = useMemo(() => findHits(doc, q, matchCase, whole), [doc, doc.version, q, matchCase, whole]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => input.current?.focus(), []);
  useEffect(() => setShowRep(mode === 'replace'), [mode]);
  const go = (i: number) => {
    if (!hits.length) return;
    const n = ((i % hits.length) + hits.length) % hits.length;
    setCur(n);
    const h = hits[n];
    doc.editing = null;
    doc.setSel({ slide: h.slide, els: [h.el], slides: [h.slide] });
  };
  const replaceHit = (h: Hit) => {
    doc.updateSlide(
      'Replace',
      (s) => {
        const fix = (e: El): El => {
          if (e.id !== h.el) return e.type === 'group' ? { ...e, children: e.children.map(fix) } : e;
          if (e.type === 'shape' && e.text && !h.cell) return { ...e, text: { ...e.text, paras: e.text.paras.map((p, i) => (i === h.para ? replaceInPara(p, h.start, h.len, rep) : p)) } };
          if (e.type === 'table' && h.cell) {
            const [r, c] = h.cell;
            return { ...e, rows: e.rows.map((row, ri) => (ri !== r ? row : { ...row, cells: row.cells.map((cell, ci) => (ci !== c ? cell : { ...cell, text: { ...cell.text, paras: cell.text.paras.map((p, i) => (i === h.para ? replaceInPara(p, h.start, h.len, rep) : p)) } })) })) };
          }
          return e;
        };
        return { ...s, elements: s.elements.map(fix) };
      },
      h.slide,
    );
  };
  const replaceAll = () => {
    // replace from the end so earlier offsets stay valid
    const all = [...hits].reverse();
    for (const h of all) replaceHit(h);
  };
  return (
    <div className="sl-find" onKeyDown={(e) => e.key === 'Escape' && onClose()}>
      <div className="sl-find-row">
        <button className="icon-btn icon-btn-sm" aria-label="Toggle replace" onClick={() => setShowRep((v) => !v)}>
          {showRep ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
        </button>
        <div className="input-icon grow">
          <Search size={14} />
          <input ref={input} className="input input-sm" placeholder="Find in presentation" value={q} onChange={(e) => (setQ(e.target.value), setCur(0))} onKeyDown={(e) => e.key === 'Enter' && go(e.shiftKey ? cur - 1 : cur + (hits.length && doc.sel.els.includes(hits[cur]?.el) ? 1 : 0))} />
        </div>
        <span className="muted small nowrap">{q ? (hits.length ? `${Math.min(cur + 1, hits.length)} of ${hits.length}` : 'No results') : ''}</span>
        <button className="icon-btn icon-btn-sm" aria-label="Previous" onClick={() => go(cur - 1)}>
          <ArrowUp size={14} />
        </button>
        <button className="icon-btn icon-btn-sm" aria-label="Next" onClick={() => go(cur + 1)}>
          <ArrowDown size={14} />
        </button>
        <button className="icon-btn icon-btn-sm" aria-label="Close" onClick={onClose}>
          <X size={14} />
        </button>
      </div>
      {showRep && (
        <div className="sl-find-row">
          <span style={{ width: 26 }} />
          <div className="input-icon grow">
            <Replace size={14} />
            <input className="input input-sm" placeholder="Replace with" value={rep} onChange={(e) => setRep(e.target.value)} />
          </div>
          <button className="btn btn-sm" disabled={!hits.length} onClick={() => hits[cur] && replaceHit(hits[cur])}>
            Replace
          </button>
          <button className="btn btn-sm" disabled={!hits.length} onClick={replaceAll}>
            All
          </button>
        </div>
      )}
      <div className="sl-find-row opts">
        <Checkbox checked={matchCase} onChange={setMatchCase} label="Match case" />
        <Checkbox checked={whole} onChange={setWhole} label="Whole words" />
      </div>
    </div>
  );
}

/* ================================================================ sorter */

export function SlideSorter({ doc, onOpen, onContextMenu }: { doc: SlidesDoc; onOpen: (i: number) => void; onContextMenu: (e: React.MouseEvent, i: number) => void }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  const [dropAt, setDropAt] = useState<number | null>(null);
  const drag = useRef<{ from: number; x: number; y: number; active: boolean } | null>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const pres = doc.pres;
  const indexAt = (x: number, y: number) => {
    const items = Array.from(gridRef.current?.querySelectorAll<HTMLElement>('[data-index]') ?? []);
    for (const it of items) {
      const r = it.getBoundingClientRect();
      if (y < r.bottom && x < r.left + r.width / 2 && y > r.top - 10) return Number(it.dataset.index);
    }
    return items.length;
  };
  return (
    <div className="sl-sorter thin-scroll" ref={gridRef}>
      {pres.slides.map((s, i) => (
        <div
          key={s.id}
          data-index={i}
          className={`sl-sorter-item${doc.sel.slides.includes(i) ? ' selected' : ''}${s.hidden ? ' is-hidden' : ''}${dropAt === i ? ' drop-before' : ''}`}
          onPointerDown={(e) => {
            if (e.button !== 0) return;
            if (e.shiftKey) {
              const a = Math.min(doc.sel.slide, i);
              const b = Math.max(doc.sel.slide, i);
              doc.setSel({ slide: i, slides: Array.from({ length: b - a + 1 }, (_, k) => a + k) });
            } else if (e.ctrlKey || e.metaKey) doc.setSel({ slide: i, slides: doc.sel.slides.includes(i) ? doc.sel.slides.filter((x) => x !== i) : [...doc.sel.slides, i] });
            else if (!doc.sel.slides.includes(i)) doc.setSel({ slide: i, slides: [i] });
            drag.current = { from: i, x: e.clientX, y: e.clientY, active: false };
            (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
          }}
          onPointerMove={(e) => {
            const d = drag.current;
            if (!d) return;
            if (!d.active && Math.hypot(e.clientX - d.x, e.clientY - d.y) < 6) return;
            d.active = true;
            setDropAt(indexAt(e.clientX, e.clientY));
          }}
          onPointerUp={(e) => {
            const d = drag.current;
            drag.current = null;
            setDropAt(null);
            if (d?.active) moveSlides(doc, doc.sel.slides.includes(d.from) ? doc.sel.slides : [d.from], indexAt(e.clientX, e.clientY));
          }}
          onDoubleClick={() => onOpen(i)}
          onContextMenu={(e) => {
            e.preventDefault();
            if (!doc.sel.slides.includes(i)) doc.setSel({ slide: i, slides: [i] });
            onContextMenu(e, i);
          }}
        >
          <Thumb pres={pres} slide={s} index={i} width={240} />
          <div className="sl-sorter-num">
            {i + 1}
            {s.transition?.type && s.transition.type !== 'none' && <Sparkles size={11} />}
          </div>
        </div>
      ))}
    </div>
  );
}
