import { ChevronDown, ChevronUp, PanelTopClose, PanelTopOpen } from 'lucide-react';
import { useEffect, useRef, useState, type ReactNode } from 'react';
import { useSettings } from '@/app/settings';
import { Menu, type MenuItem } from './menu';
import { Popover } from './popover';

export interface RibbonTab {
  id: string;
  label: string;
  content: ReactNode;
  /** Contextual tabs (e.g. "Table", "Picture") get the accent style. */
  contextual?: boolean;
  hidden?: boolean;
}

export function Ribbon({
  app,
  tabs,
  active,
  onActive,
  onFile,
  right,
  fileLabel = 'File',
}: {
  app: 'doc' | 'sheet' | 'slides';
  tabs: RibbonTab[];
  active: string;
  onActive: (id: string) => void;
  onFile: () => void;
  right?: ReactNode;
  fileLabel?: string;
}) {
  const mode = useSettings((s) => s.settings.ribbonMode);
  const setSettings = useSettings((s) => s.update);
  const [collapsed, setCollapsed] = useState(false);
  const [peek, setPeek] = useState(false);
  const bodyRef = useRef<HTMLDivElement>(null);
  const visibleTabs = tabs.filter((t) => !t.hidden);
  const current = visibleTabs.find((t) => t.id === active) ?? visibleTabs[0];

  useEffect(() => {
    const el = bodyRef.current;
    if (!el) return;
    // Mouse wheel scrolls an overflowing ribbon horizontally
    const onWheel = (e: WheelEvent) => {
      if (el.scrollWidth > el.clientWidth && Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
        el.scrollLeft += e.deltaY;
        e.preventDefault();
      }
    };
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, []);

  const showBody = !collapsed || peek;

  return (
    <div className={`ribbon app-${app} ribbon-${mode}${collapsed ? ' collapsed' : ''}${peek ? ' peek' : ''}`}>
      <div className="ribbon-tabs" role="tablist">
        <button type="button" className="ribbon-file" onClick={onFile}>
          {fileLabel}
        </button>
        {visibleTabs.map((t) => (
          <button
            type="button"
            role="tab"
            aria-selected={t.id === current?.id}
            key={t.id}
            className={`ribbon-tab${t.id === current?.id ? ' active' : ''}${t.contextual ? ' contextual' : ''}`}
            onClick={() => {
              if (collapsed) {
                if (t.id === current?.id) setPeek((p) => !p);
                else setPeek(true);
              }
              onActive(t.id);
            }}
            onDoubleClick={() => {
              setCollapsed((c) => !c);
              setPeek(false);
            }}
          >
            {t.label}
          </button>
        ))}
        <span className="spacer" />
        {right}
        <button
          type="button"
          className="icon-btn icon-btn-sm ribbon-mode"
          data-tip={mode === 'full' ? 'Switch to simplified ribbon' : 'Switch to classic ribbon'}
          onClick={() => setSettings({ ribbonMode: mode === 'full' ? 'simplified' : 'full' })}
        >
          {mode === 'full' ? <PanelTopClose size={15} /> : <PanelTopOpen size={15} />}
        </button>
        <button
          type="button"
          className="icon-btn icon-btn-sm"
          data-tip={collapsed ? 'Always show the ribbon' : 'Collapse the ribbon (show tabs only)'}
          data-tip-key="Ctrl+F1"
          onClick={() => {
            setCollapsed((c) => !c);
            setPeek(false);
          }}
        >
          {collapsed ? <ChevronDown size={15} /> : <ChevronUp size={15} />}
        </button>
      </div>
      {showBody && (
        <div
          className="ribbon-body thin-scroll"
          ref={bodyRef}
          onMouseLeave={() => {
            if (peek) setTimeout(() => setPeek(false), 400);
          }}
        >
          <div className="ribbon-panel" key={current?.id}>
            {current?.content}
          </div>
        </div>
      )}
    </div>
  );
}

export function RibbonGroup({ label, children, className = '' }: { label: string; children: ReactNode; className?: string }) {
  return (
    <div className={`rgroup ${className}`}>
      <div className="rgroup-content">{children}</div>
      <div className="rgroup-label">{label}</div>
    </div>
  );
}

export function RRows({ children }: { children: ReactNode }) {
  return <div className="rrows">{children}</div>;
}

export function RRow({ children }: { children: ReactNode }) {
  return <div className="rrow">{children}</div>;
}

export function RSep() {
  return <span className="rsep" aria-hidden />;
}

export interface RButtonProps {
  icon?: ReactNode;
  label?: string;
  tip?: string;
  keys?: string;
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
  /** Show the text label next to the icon */
  showLabel?: boolean;
  className?: string;
}

/** Small ribbon button (icon, optional label). Keeps editor focus by not stealing mousedown. */
export function RButton({ icon, label, tip, keys, onClick, active, disabled, showLabel, className = '' }: RButtonProps) {
  return (
    <button
      type="button"
      className={`rbtn${active ? ' active' : ''}${showLabel ? ' with-label' : ''} ${className}`}
      data-tip={tip ?? label}
      data-tip-key={keys}
      aria-label={label ?? tip}
      aria-pressed={active === undefined ? undefined : active}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {icon}
      {showLabel && label && <span className="rbtn-label">{label}</span>}
    </button>
  );
}

/** Large ribbon button: icon above label. Optional dropdown menu. */
export function RBigButton({
  icon,
  label,
  tip,
  keys,
  onClick,
  menu,
  active,
  disabled,
  accent,
}: {
  icon: ReactNode;
  label: string;
  tip?: string;
  keys?: string;
  onClick?: () => void;
  menu?: MenuItem[] | (() => MenuItem[]);
  active?: boolean;
  disabled?: boolean;
  accent?: boolean;
}) {
  const ref = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const items = open ? (typeof menu === 'function' ? menu() : menu) : undefined;
  return (
    <>
      <button
        ref={ref}
        type="button"
        className={`rbig${active ? ' active' : ''}${accent ? ' accent' : ''}${open ? ' open' : ''}`}
        data-tip={tip}
        data-tip-key={keys}
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => (menu && !onClick ? setOpen((o) => !o) : onClick?.())}
      >
        <span className="rbig-icon">{icon}</span>
        <span className="rbig-label">
          {label}
          {menu && !onClick && <ChevronDown size={11} className="rbig-caret" />}
        </span>
      </button>
      {menu && items && <Menu open={open} anchor={ref.current} items={items} onClose={() => setOpen(false)} ignore={[ref.current]} />}
    </>
  );
}

/** Button that opens a dropdown menu or custom panel. */
export function RDropdown({
  icon,
  label,
  tip,
  menu,
  panel,
  showLabel = true,
  width,
  className = '',
  active,
}: {
  icon?: ReactNode;
  label?: string;
  tip?: string;
  menu?: MenuItem[] | (() => MenuItem[]);
  panel?: (close: () => void) => ReactNode;
  showLabel?: boolean;
  width?: number;
  className?: string;
  active?: boolean;
}) {
  const ref = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const items = open && menu ? (typeof menu === 'function' ? menu() : menu) : undefined;
  return (
    <>
      <button
        ref={ref}
        type="button"
        className={`rbtn rdrop${showLabel && label ? ' with-label' : ''}${open ? ' open' : ''}${active ? ' active' : ''} ${className}`}
        data-tip={tip ?? label}
        style={width ? { width } : undefined}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((o) => !o)}
      >
        {icon}
        {showLabel && label && <span className="rbtn-label">{label}</span>}
        <ChevronDown size={11} className="rdrop-caret" />
      </button>
      {items && <Menu open={open} anchor={ref.current} items={items} onClose={() => setOpen(false)} ignore={[ref.current]} />}
      {panel && (
        <Popover open={open} anchor={ref.current} onClose={() => setOpen(false)} ignore={[ref.current]}>
          {open && panel(() => setOpen(false))}
        </Popover>
      )}
    </>
  );
}

/** Split button: main action + dropdown with alternatives. */
export function RSplit({
  icon,
  label,
  tip,
  keys,
  onClick,
  menu,
  big,
  active,
}: {
  icon: ReactNode;
  label?: string;
  tip?: string;
  keys?: string;
  onClick: () => void;
  menu: MenuItem[] | (() => MenuItem[]);
  big?: boolean;
  active?: boolean;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);
  const items = open ? (typeof menu === 'function' ? menu() : menu) : undefined;
  return (
    <div ref={ref} className={`split-btn${big ? ' big' : ''}${open ? ' open' : ''}${active ? ' active' : ''}`}>
      <button type="button" className="split-main" data-tip={tip ?? label} data-tip-key={keys} onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
        {big ? <span className="rbig-icon">{icon}</span> : icon}
        {label && <span className="split-label">{label}</span>}
      </button>
      <button type="button" className="split-arrow" aria-label={`${label ?? tip} options`} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen((o) => !o)}>
        <ChevronDown size={11} />
      </button>
      {items && <Menu open={open} anchor={ref.current} items={items} onClose={() => setOpen(false)} ignore={[ref.current]} />}
    </div>
  );
}
