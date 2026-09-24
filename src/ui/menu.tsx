import { ChevronRight, Check } from 'lucide-react';
import { createContext, useCallback, useContext, useEffect, useRef, useState, type ReactNode } from 'react';
import { create } from 'zustand';
import { Popover, type Anchor, type Placement } from './popover';

export interface MenuItem {
  key?: string;
  label?: ReactNode;
  icon?: ReactNode;
  shortcut?: string;
  hint?: string;
  onSelect?: () => void;
  disabled?: boolean;
  checked?: boolean;
  danger?: boolean;
  submenu?: MenuItem[] | (() => MenuItem[]);
  separator?: boolean;
  header?: string;
  /** Arbitrary content rendered in place of a normal row (e.g. a colour grid). */
  custom?: ReactNode | ((close: () => void) => ReactNode);
  keepOpen?: boolean;
}

const DoneCtx = createContext<() => void>(() => undefined);

export interface MenuProps {
  open: boolean;
  anchor: Anchor;
  items: MenuItem[];
  onClose: () => void;
  placement?: Placement;
  ignore?: Array<HTMLElement | null | undefined>;
  className?: string;
  minWidth?: number;
}

export function Menu({ open, anchor, items, onClose, placement = 'bottom-start', ignore, className = '', minWidth }: MenuProps) {
  const parentDone = useContext(DoneCtx);
  const isNested = useContext(NestedCtx);
  const done = isNested ? parentDone : onClose;
  return (
    <Popover open={open} anchor={anchor} onClose={onClose} placement={placement} ignore={ignore} role="menu" zIndex={isNested ? 1101 : 1100}>
      <DoneCtx.Provider value={done}>
        <NestedCtx.Provider value={true}>
          <MenuList items={items} onClose={onClose} className={className} minWidth={minWidth} />
        </NestedCtx.Provider>
      </DoneCtx.Provider>
    </Popover>
  );
}

const NestedCtx = createContext(false);

function resolveItems(items: MenuItem[] | (() => MenuItem[]) | undefined): MenuItem[] {
  if (!items) return [];
  return typeof items === 'function' ? items() : items;
}

export function MenuList({ items, className = '', minWidth }: { items: MenuItem[]; onClose?: () => void; className?: string; minWidth?: number }) {
  const done = useContext(DoneCtx);
  const [active, setActive] = useState(-1);
  const [subOpen, setSubOpen] = useState(-1);
  const rowRefs = useRef<Array<HTMLDivElement | null>>([]);
  const listRef = useRef<HTMLDivElement>(null);
  const hoverTimer = useRef<ReturnType<typeof setTimeout>>(undefined);

  const selectable = useCallback((i: number) => {
    const it = items[i];
    return it && !it.separator && !it.header && !it.custom && !it.disabled;
  }, [items]);

  useEffect(() => {
    listRef.current?.focus({ preventScroll: true });
  }, []);

  useEffect(() => () => clearTimeout(hoverTimer.current), []);

  const activate = (i: number) => {
    const it = items[i];
    if (!it || it.disabled) return;
    if (it.submenu) {
      setSubOpen(i);
      return;
    }
    it.onSelect?.();
    if (!it.keepOpen) done();
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (subOpen >= 0 && e.key !== 'ArrowLeft') return;
    const n = items.length;
    const step = (dir: number) => {
      let i = active;
      for (let k = 0; k < n; k++) {
        i = (i + dir + n) % n;
        if (selectable(i)) {
          setActive(i);
          rowRefs.current[i]?.scrollIntoView({ block: 'nearest' });
          return;
        }
      }
    };
    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        step(1);
        break;
      case 'ArrowUp':
        e.preventDefault();
        step(-1);
        break;
      case 'Home':
        e.preventDefault();
        setActive(-1);
        step(1);
        break;
      case 'End':
        e.preventDefault();
        setActive(n);
        step(-1);
        break;
      case 'Enter':
      case ' ':
        e.preventDefault();
        if (active >= 0) activate(active);
        break;
      case 'ArrowRight':
        if (active >= 0 && items[active]?.submenu) {
          e.preventDefault();
          setSubOpen(active);
        }
        break;
      case 'ArrowLeft':
        if (subOpen >= 0) {
          e.preventDefault();
          setSubOpen(-1);
          listRef.current?.focus();
        }
        break;
      default:
        // type-ahead: jump to first item starting with the typed letter
        if (e.key.length === 1 && /\S/.test(e.key)) {
          const ch = e.key.toLowerCase();
          const start = active + 1;
          for (let k = 0; k < n; k++) {
            const i = (start + k) % n;
            const label = items[i]?.label;
            if (selectable(i) && typeof label === 'string' && label.toLowerCase().startsWith(ch)) {
              setActive(i);
              rowRefs.current[i]?.scrollIntoView({ block: 'nearest' });
              break;
            }
          }
        }
    }
  };

  return (
    <div
      ref={listRef}
      className={`menu ${className}`}
      style={minWidth ? { minWidth } : undefined}
      tabIndex={-1}
      onKeyDown={onKeyDown}
      onMouseLeave={() => subOpen < 0 && setActive(-1)}
    >
      {items.map((it, i) => {
        if (it.separator) return <div key={it.key ?? `sep${i}`} className="menu-sep" role="separator" />;
        if (it.header) return <div key={it.key ?? `h${i}`} className="menu-header">{it.header}</div>;
        if (it.custom) {
          return (
            <div key={it.key ?? `c${i}`} className="menu-custom">
              {typeof it.custom === 'function' ? it.custom(done) : it.custom}
            </div>
          );
        }
        const isActive = active === i;
        return (
          <div
            key={it.key ?? `${i}`}
            ref={(el) => {
              rowRefs.current[i] = el;
            }}
            role="menuitem"
            aria-disabled={it.disabled || undefined}
            aria-checked={it.checked === undefined ? undefined : it.checked}
            className={`menu-item${isActive ? ' active' : ''}${it.disabled ? ' disabled' : ''}${it.danger ? ' danger' : ''}${subOpen === i ? ' sub-open' : ''}`}
            onMouseEnter={() => {
              setActive(i);
              clearTimeout(hoverTimer.current);
              if (it.submenu && !it.disabled) hoverTimer.current = setTimeout(() => setSubOpen(i), 140);
              else if (subOpen >= 0) hoverTimer.current = setTimeout(() => setSubOpen(-1), 200);
            }}
            onClick={(e) => {
              e.stopPropagation();
              activate(i);
            }}
          >
            <span className="menu-check">{it.checked ? <Check size={14} strokeWidth={2.4} /> : null}</span>
            {it.icon !== undefined && <span className="menu-icon">{it.icon}</span>}
            <span className="menu-label">
              {it.label}
              {it.hint && <span className="menu-hint">{it.hint}</span>}
            </span>
            {it.shortcut && <span className="menu-shortcut">{it.shortcut}</span>}
            {it.submenu && <ChevronRight size={14} className="menu-arrow" />}
            {it.submenu && subOpen === i && (
              <Menu
                open
                anchor={rowRefs.current[i]}
                items={resolveItems(it.submenu)}
                placement="right-start"
                onClose={() => {
                  setSubOpen(-1);
                  listRef.current?.focus();
                }}
              />
            )}
          </div>
        );
      })}
    </div>
  );
}

/* --------------------------------------------------------------- dropdown */

export function useMenuToggle() {
  const [anchor, setAnchor] = useState<HTMLElement | null>(null);
  return {
    anchor,
    open: Boolean(anchor),
    toggle: (el: HTMLElement) => setAnchor((a) => (a ? null : el)),
    close: () => setAnchor(null),
  };
}

/* ----------------------------------------------------------- context menu */

interface CtxState {
  point: { x: number; y: number } | null;
  items: MenuItem[];
  show(point: { x: number; y: number }, items: MenuItem[]): void;
  hide(): void;
}

export const useContextMenuStore = create<CtxState>((set) => ({
  point: null,
  items: [],
  show: (point, items) => set({ point, items }),
  hide: () => set({ point: null, items: [] }),
}));

export function openContextMenu(e: { clientX: number; clientY: number; preventDefault?: () => void } | { x: number; y: number }, items: MenuItem[]) {
  if ('preventDefault' in e && e.preventDefault) e.preventDefault();
  const point = 'clientX' in e ? { x: e.clientX, y: e.clientY } : e;
  const filtered = items.filter((it, i, arr) => !(it.separator && (i === 0 || arr[i - 1]?.separator || i === arr.length - 1)));
  useContextMenuStore.getState().show(point, filtered);
}

export function ContextMenuHost() {
  const { point, items, hide } = useContextMenuStore();
  return <Menu open={Boolean(point)} anchor={point} items={items} onClose={hide} placement="bottom-start" />;
}

export function MenuButtonContent({ children, icon }: { children?: ReactNode; icon?: ReactNode }) {
  return (
    <>
      {icon}
      {children}
    </>
  );
}
