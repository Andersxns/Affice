import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { computePosition } from './popover';

/**
 * One tooltip for the whole app. Any element with `data-tip="…"` (and optional
 * `data-tip-key="Ctrl+B"`) gets a tooltip after a short hover delay.
 */
export function TooltipHost() {
  const [tip, setTip] = useState<{ text: string; keys?: string; rect: DOMRect } | null>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const timer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const current = useRef<HTMLElement | null>(null);
  const warm = useRef(0);

  useEffect(() => {
    const show = (el: HTMLElement) => {
      const text = el.getAttribute('data-tip');
      if (!text) return;
      setTip({ text, keys: el.getAttribute('data-tip-key') ?? undefined, rect: el.getBoundingClientRect() });
    };
    const over = (e: PointerEvent) => {
      if (e.pointerType === 'touch') return;
      const el = (e.target as HTMLElement | null)?.closest?.('[data-tip]') as HTMLElement | null;
      if (el === current.current) return;
      current.current = el;
      clearTimeout(timer.current);
      if (!el) {
        setTip(null);
        return;
      }
      const recentlyShown = Date.now() - warm.current < 600;
      timer.current = setTimeout(() => show(el), recentlyShown ? 60 : 520);
    };
    const hide = () => {
      clearTimeout(timer.current);
      if (current.current) warm.current = Date.now();
      current.current = null;
      setTip(null);
    };
    document.addEventListener('pointerover', over);
    document.addEventListener('pointerdown', hide, true);
    document.addEventListener('keydown', hide, true);
    document.addEventListener('wheel', hide, { passive: true, capture: true });
    window.addEventListener('blur', hide);
    return () => {
      document.removeEventListener('pointerover', over);
      document.removeEventListener('pointerdown', hide, true);
      document.removeEventListener('keydown', hide, true);
      document.removeEventListener('wheel', hide, true);
      window.removeEventListener('blur', hide);
    };
  }, []);

  useEffect(() => {
    if (!tip || !ref.current) {
      setPos(null);
      return;
    }
    const r = ref.current.getBoundingClientRect();
    const p = computePosition(tip.rect, r.width, r.height, 'bottom', 8);
    setPos({ left: p.left, top: p.top });
  }, [tip]);

  if (!tip) return null;
  return createPortal(
    <div
      ref={ref}
      className="tooltip"
      role="tooltip"
      style={{ left: pos?.left ?? -9999, top: pos?.top ?? -9999, visibility: pos ? 'visible' : 'hidden' }}
    >
      <span>{tip.text}</span>
      {tip.keys && <span className="tooltip-keys">{tip.keys}</span>}
    </div>,
    document.body,
  );
}
