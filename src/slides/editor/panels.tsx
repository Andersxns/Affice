import { EyeOff, Plus, Sparkles } from 'lucide-react';
import { memo, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import { isMod } from '@/lib/utils';
import type { SlidesDoc } from '../doc';
import type { Presentation, Slide } from '../model';
import { moveSlides, setNotes } from '../ops';
import { SlideView } from '../render/SlideView';

/* ================================================================ thumbs */

/** A slide rendered small; renders only once scrolled into view (cheap on old machines). */
export const Thumb = memo(function Thumb({ pres, slide, index, width, eager }: { pres: Presentation; slide: Slide; index: number; width: number; eager?: boolean }) {
  const ref = useRef<HTMLDivElement>(null);
  const [visible, setVisible] = useState(!!eager);
  useEffect(() => {
    if (visible || !ref.current) return;
    const io = new IntersectionObserver((entries) => {
      if (entries.some((e) => e.isIntersecting)) {
        setVisible(true);
        io.disconnect();
      }
    }, { rootMargin: '300px' });
    io.observe(ref.current);
    return () => io.disconnect();
  }, [visible]);
  const scale = width / pres.size.w;
  const height = Math.round(pres.size.h * scale);
  return (
    <div ref={ref} className="sl-thumb-view" style={{ width, height }}>
      {visible && (
        <div style={{ transform: `scale(${scale})`, transformOrigin: '0 0', width: pres.size.w, height: pres.size.h }}>
          <SlideView pres={pres} slide={slide} slideNumber={index + 1} mode="thumb" />
        </div>
      )}
    </div>
  );
});

interface PanelProps {
  doc: SlidesDoc;
  onContextMenu: (e: React.MouseEvent, index: number | null) => void;
  onAdd: () => void;
  onKey: (e: React.KeyboardEvent) => void;
  width: number;
}

export function SlidePanel({ doc, onContextMenu, onAdd, onKey, width }: PanelProps) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  const listRef = useRef<HTMLDivElement>(null);
  const drag = useRef<{ startY: number; from: number; active: boolean; pointer: number } | null>(null);
  const [dropAt, setDropAt] = useState<number | null>(null);
  const pres = doc.pres;
  const thumbW = Math.max(80, width - 44);

  // keep the current slide in view
  useLayoutEffect(() => {
    const el = listRef.current?.querySelector<HTMLElement>(`[data-index="${doc.sel.slide}"]`);
    el?.scrollIntoView({ block: 'nearest' });
  }, [doc.sel.slide]);

  const select = (i: number, e: React.MouseEvent | React.PointerEvent) => {
    const cur = doc.sel;
    if (e.shiftKey) {
      const a = Math.min(cur.slide, i);
      const b = Math.max(cur.slide, i);
      doc.setSel({ slide: i, slides: Array.from({ length: b - a + 1 }, (_, k) => a + k) });
    } else if (isMod(e)) {
      const has = cur.slides.includes(i);
      const slides = has && cur.slides.length > 1 ? cur.slides.filter((x) => x !== i) : [...cur.slides, i];
      doc.setSel({ slide: has && cur.slides.length > 1 ? slides[0] : i, slides });
    } else if (!cur.slides.includes(i) || cur.slides.length === 1) {
      doc.editing = null;
      doc.setSel({ slide: i, slides: [i] });
    } else doc.setSel({ slide: i, slides: cur.slides });
  };

  const indexAt = (clientY: number): number => {
    const items = Array.from(listRef.current?.querySelectorAll<HTMLElement>('[data-index]') ?? []);
    for (const it of items) {
      const r = it.getBoundingClientRect();
      if (clientY < r.top + r.height / 2) return Number(it.dataset.index);
    }
    return items.length;
  };

  return (
    <div
      className="sl-panel thin-scroll"
      ref={listRef}
      tabIndex={0}
      onKeyDown={onKey}
      onContextMenu={(e) => {
        if (!(e.target as HTMLElement).closest('[data-index]')) {
          e.preventDefault();
          onContextMenu(e, null);
        }
      }}
    >
      {pres.slides.map((s, i) => {
        const selected = doc.sel.slides.includes(i);
        return (
          <div key={s.id}>
            {dropAt === i && <div className="sl-drop-line" />}
            <div
              className={`sl-thumb${i === doc.sel.slide ? ' current' : ''}${selected ? ' selected' : ''}${s.hidden ? ' is-hidden' : ''}`}
              data-index={i}
              onPointerDown={(e) => {
                if (e.button !== 0) return;
                if (!doc.sel.slides.includes(i) || e.shiftKey || isMod(e)) select(i, e);
                drag.current = { startY: e.clientY, from: i, active: false, pointer: e.pointerId };
                (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
              }}
              onPointerMove={(e) => {
                const d = drag.current;
                if (!d) return;
                if (!d.active && Math.abs(e.clientY - d.startY) < 6) return;
                d.active = true;
                setDropAt(indexAt(e.clientY));
              }}
              onPointerUp={(e) => {
                const d = drag.current;
                drag.current = null;
                if (d?.active) {
                  const to = indexAt(e.clientY);
                  setDropAt(null);
                  moveSlides(doc, doc.sel.slides.includes(d.from) ? doc.sel.slides : [d.from], to);
                } else if (d && !e.shiftKey && !isMod(e)) select(i, e);
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                if (!doc.sel.slides.includes(i)) select(i, e);
                onContextMenu(e, i);
              }}
            >
              <div className="sl-thumb-num">
                <span>{i + 1}</span>
                {s.hidden && <EyeOff size={11} />}
                {(s.transition?.type && s.transition.type !== 'none') || s.anims?.length ? <Sparkles size={11} className="sl-thumb-fx" /> : null}
              </div>
              <div className="sl-thumb-frame">
                <Thumb pres={pres} slide={s} index={i} width={thumbW} eager={Math.abs(i - doc.sel.slide) < 8} />
              </div>
            </div>
          </div>
        );
      })}
      {dropAt === pres.slides.length && <div className="sl-drop-line" />}
      <button type="button" className="sl-add-slide" onClick={onAdd} data-tip="New slide (Ctrl+M)">
        <Plus size={16} /> New slide
      </button>
    </div>
  );
}

/* ================================================================= notes */

export function NotesPane({ doc, height, onResize }: { doc: SlidesDoc; height: number; onResize: (h: number) => void }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  const slide = doc.slide;
  const [text, setText] = useState(slide?.notes ?? '');
  const slideId = slide?.id;
  useEffect(() => {
    setText(doc.slide?.notes ?? '');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [slideId, doc.version]);
  const startResize = (e: React.PointerEvent) => {
    const y0 = e.clientY;
    const h0 = height;
    const move = (ev: PointerEvent) => onResize(Math.max(48, Math.min(480, h0 - (ev.clientY - y0))));
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  return (
    <div className="sl-notes" style={{ height }}>
      <div className="sl-notes-grip" onPointerDown={startResize} />
      <textarea
        className="sl-notes-input thin-scroll"
        placeholder="Click to add speaker notes"
        value={text}
        onChange={(e) => {
          setText(e.target.value);
          setNotes(doc, e.target.value);
        }}
        spellCheck
      />
    </div>
  );
}
