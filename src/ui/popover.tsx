import { createContext, useContext, useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

export type Placement = 'bottom-start' | 'bottom-end' | 'bottom' | 'top-start' | 'top-end' | 'top' | 'right-start' | 'left-start' | 'right-end';

export type Anchor = HTMLElement | DOMRect | { x: number; y: number } | null | undefined;

function anchorRect(anchor: Anchor): DOMRect | null {
  if (!anchor) return null;
  if (anchor instanceof HTMLElement) return anchor.getBoundingClientRect();
  if (anchor instanceof DOMRect) return anchor;
  return new DOMRect(anchor.x, anchor.y, 0, 0);
}

export function computePosition(a: DOMRect, w: number, h: number, placement: Placement, offset: number, margin = 6) {
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  let [side, align] = placement.split('-') as [string, string | undefined];
  let left = 0;
  let top = 0;

  const fitsBelow = a.bottom + offset + h <= vh - margin;
  const fitsAbove = a.top - offset - h >= margin;
  const fitsRight = a.right + offset + w <= vw - margin;
  const fitsLeft = a.left - offset - w >= margin;
  if (side === 'bottom' && !fitsBelow && fitsAbove) side = 'top';
  else if (side === 'top' && !fitsAbove && fitsBelow) side = 'bottom';
  else if (side === 'right' && !fitsRight && fitsLeft) side = 'left';
  else if (side === 'left' && !fitsLeft && fitsRight) side = 'right';

  if (side === 'bottom' || side === 'top') {
    top = side === 'bottom' ? a.bottom + offset : a.top - offset - h;
    if (align === 'start') left = a.left;
    else if (align === 'end') left = a.right - w;
    else left = a.left + a.width / 2 - w / 2;
  } else {
    left = side === 'right' ? a.right + offset : a.left - offset - w;
    if (align === 'end') top = a.bottom - h;
    else top = a.top - (align === 'start' ? 4 : 0);
  }
  left = Math.max(margin, Math.min(left, vw - w - margin));
  top = Math.max(margin, Math.min(top, vh - h - margin));
  return { left, top, side };
}

const RootCtx = createContext<string | null>(null);
export const usePopoverRoot = () => useContext(RootCtx);

let popoverSeq = 0;

export interface PopoverProps {
  open: boolean;
  anchor: Anchor;
  onClose: () => void;
  placement?: Placement;
  offset?: number;
  className?: string;
  style?: CSSProperties;
  children: ReactNode;
  /** Elements (e.g. the toggle button) whose clicks should not count as "outside". */
  ignore?: Array<HTMLElement | null | undefined>;
  closeOnEscape?: boolean;
  role?: string;
  zIndex?: number;
  autoFocus?: boolean;
  matchWidth?: boolean;
}

/** A positioned floating panel rendered in a portal, with outside-click and Escape handling. */
export function Popover({
  open,
  anchor,
  onClose,
  placement = 'bottom-start',
  offset = 4,
  className = '',
  style,
  children,
  ignore,
  closeOnEscape = true,
  role,
  zIndex = 1000,
  autoFocus = false,
  matchWidth = false,
}: PopoverProps) {
  const parentRoot = usePopoverRoot();
  const [rootId] = useState(() => parentRoot ?? `pop${++popoverSeq}`);
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number; side: string } | null>(null);
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  const ignoreRef = useRef(ignore);
  ignoreRef.current = ignore;

  useLayoutEffect(() => {
    if (!open) {
      setPos(null);
      return;
    }
    const el = ref.current;
    const a = anchorRect(anchor);
    if (!el || !a) return;
    if (matchWidth) el.style.minWidth = `${a.width}px`;
    const r = el.getBoundingClientRect();
    setPos(computePosition(a, r.width, r.height, placement, offset));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, anchor, placement, offset]);

  // Reposition when content size changes (e.g. async lists)
  useEffect(() => {
    if (!open || !ref.current || typeof ResizeObserver === 'undefined') return;
    const el = ref.current;
    const ro = new ResizeObserver(() => {
      const a = anchorRect(anchor);
      if (!a) return;
      const r = el.getBoundingClientRect();
      setPos(computePosition(a, r.width, r.height, placement, offset));
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [open, anchor, placement, offset]);

  useEffect(() => {
    if (!open) return;
    // Every popover in a group (a menu and its submenus) shares one root id; clicks
    // inside any member of the group — or on an ignored element — are not "outside".
    const onDown = (e: PointerEvent) => {
      for (const n of e.composedPath()) {
        if (!(n instanceof HTMLElement)) continue;
        if (n === ref.current || n.getAttribute('data-pop-root') === rootId) return;
        if (ignoreRef.current?.includes(n)) return;
      }
      onCloseRef.current();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && closeOnEscape) {
        e.stopPropagation();
        onCloseRef.current();
      }
    };
    const onBlur = () => onCloseRef.current();
    const t = setTimeout(() => {
      document.addEventListener('pointerdown', onDown, true);
    }, 0);
    document.addEventListener('keydown', onKey, true);
    window.addEventListener('blur', onBlur);
    window.addEventListener('resize', onBlur);
    return () => {
      clearTimeout(t);
      document.removeEventListener('pointerdown', onDown, true);
      document.removeEventListener('keydown', onKey, true);
      window.removeEventListener('blur', onBlur);
      window.removeEventListener('resize', onBlur);
    };
  }, [open, rootId, closeOnEscape]);

  useEffect(() => {
    if (open && autoFocus && pos && ref.current) {
      const f = ref.current.querySelector<HTMLElement>('input, [tabindex="0"], button:not([disabled])');
      (f ?? ref.current).focus({ preventScroll: true });
    }
  }, [open, autoFocus, pos]);

  if (!open) return null;
  const side = pos?.side ?? 'bottom';
  return createPortal(
    <RootCtx.Provider value={rootId}>
      <div
        ref={ref}
        role={role}
        tabIndex={-1}
        data-pop-root={rootId}
        className={`popover side-${side} ${className}`}
        style={{
          position: 'fixed',
          left: pos?.left ?? -9999,
          top: pos?.top ?? -9999,
          visibility: pos ? 'visible' : 'hidden',
          zIndex,
          ...style,
        }}
        onContextMenu={(e) => e.preventDefault()}
      >
        {children}
      </div>
    </RootCtx.Provider>,
    document.body,
  );
}
