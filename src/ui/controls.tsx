import { ChevronDown, ChevronUp } from 'lucide-react';
import { forwardRef, useEffect, useMemo, useRef, useState, type ButtonHTMLAttributes, type CSSProperties, type ReactNode } from 'react';
import { Popover } from './popover';

/* ----------------------------------------------------------------- buttons */

export interface IconButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  tip?: string;
  tipKey?: string;
  active?: boolean;
  size?: 'sm' | 'md' | 'lg';
}

export const IconButton = forwardRef<HTMLButtonElement, IconButtonProps>(function IconButton(
  { tip, tipKey, active, size = 'md', className = '', children, ...rest },
  ref,
) {
  return (
    <button
      ref={ref}
      type="button"
      data-tip={tip}
      data-tip-key={tipKey}
      aria-label={rest['aria-label'] ?? tip}
      aria-pressed={active === undefined ? undefined : active}
      className={`icon-btn icon-btn-${size}${active ? ' active' : ''} ${className}`}
      {...rest}
    >
      {children}
    </button>
  );
});

export function Switch({ checked, onChange, disabled, label }: { checked: boolean; onChange: (v: boolean) => void; disabled?: boolean; label?: string }) {
  return (
    <button
      type="button"
      role="switch"
      aria-checked={checked}
      aria-label={label}
      disabled={disabled}
      className={`switch${checked ? ' on' : ''}`}
      onClick={() => onChange(!checked)}
    >
      <span className="switch-knob" />
    </button>
  );
}

export function Checkbox({ checked, onChange, label, disabled }: { checked: boolean; onChange: (v: boolean) => void; label?: ReactNode; disabled?: boolean }) {
  return (
    <label className={`checkbox${disabled ? ' disabled' : ''}`}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => onChange(e.target.checked)} />
      <span className="checkbox-box" aria-hidden />
      {label && <span>{label}</span>}
    </label>
  );
}

export function Segmented<T extends string>({
  value,
  options,
  onChange,
  size = 'md',
}: {
  value: T;
  options: Array<{ value: T; label: ReactNode; tip?: string }>;
  onChange: (v: T) => void;
  size?: 'sm' | 'md';
}) {
  const idx = Math.max(0, options.findIndex((o) => o.value === value));
  return (
    <div className={`segmented segmented-${size}`} role="radiogroup" style={{ ['--seg-count' as string]: options.length, ['--seg-index' as string]: idx } as CSSProperties}>
      <span className="segmented-thumb" aria-hidden />
      {options.map((o) => (
        <button
          key={o.value}
          type="button"
          role="radio"
          aria-checked={o.value === value}
          data-tip={o.tip}
          className={`segmented-opt${o.value === value ? ' on' : ''}`}
          onClick={() => onChange(o.value)}
        >
          {o.label}
        </button>
      ))}
    </div>
  );
}

export function Slider({
  value,
  min,
  max,
  step = 1,
  onChange,
  width = 120,
  label,
}: {
  value: number;
  min: number;
  max: number;
  step?: number;
  onChange: (v: number) => void;
  width?: number;
  label?: string;
}) {
  const pct = ((value - min) / (max - min)) * 100;
  return (
    <input
      type="range"
      className="slider"
      aria-label={label}
      min={min}
      max={max}
      step={step}
      value={value}
      style={{ width, ['--pct' as string]: `${pct}%` } as CSSProperties}
      onChange={(e) => onChange(Number(e.target.value))}
    />
  );
}

/* ---------------------------------------------------------- number spinner */

export function NumberField({
  value,
  onChange,
  min = -Infinity,
  max = Infinity,
  step = 1,
  width = 72,
  suffix,
  precision = 2,
  disabled,
  label,
}: {
  value: number | null;
  onChange: (v: number) => void;
  min?: number;
  max?: number;
  step?: number;
  width?: number;
  suffix?: string;
  precision?: number;
  disabled?: boolean;
  label?: string;
}) {
  const fmt = (v: number | null) => (v === null || Number.isNaN(v) ? '' : String(Number(v.toFixed(precision))));
  const [text, setText] = useState(fmt(value));
  const focused = useRef(false);
  useEffect(() => {
    if (!focused.current) setText(fmt(value));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);
  const commit = (t: string) => {
    const v = parseFloat(t.replace(',', '.'));
    if (Number.isFinite(v)) {
      const c = Math.min(max, Math.max(min, v));
      onChange(c);
      setText(fmt(c));
    } else setText(fmt(value));
  };
  const bump = (dir: number) => {
    const base = value ?? 0;
    const c = Math.min(max, Math.max(min, Number((base + dir * step).toFixed(precision))));
    onChange(c);
    setText(fmt(c));
  };
  return (
    <div className={`number-field${disabled ? ' disabled' : ''}`} style={{ width }}>
      <input
        className="number-input"
        aria-label={label}
        value={text}
        disabled={disabled}
        onFocus={(e) => {
          focused.current = true;
          e.currentTarget.select();
        }}
        onBlur={(e) => {
          focused.current = false;
          commit(e.currentTarget.value);
        }}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            commit(e.currentTarget.value);
            e.currentTarget.blur();
          } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            bump(1);
          } else if (e.key === 'ArrowDown') {
            e.preventDefault();
            bump(-1);
          } else if (e.key === 'Escape') {
            setText(fmt(value));
            e.currentTarget.blur();
          }
        }}
      />
      {suffix && <span className="number-suffix">{suffix}</span>}
      <div className="number-spin">
        <button type="button" tabIndex={-1} disabled={disabled} onClick={() => bump(1)} aria-label="Increase">
          <ChevronUp size={11} />
        </button>
        <button type="button" tabIndex={-1} disabled={disabled} onClick={() => bump(-1)} aria-label="Decrease">
          <ChevronDown size={11} />
        </button>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------- combobox */

export interface ComboOption {
  value: string;
  label?: string;
  group?: string;
  style?: CSSProperties;
  hint?: string;
}

export function ComboBox({
  value,
  options,
  onCommit,
  width = 150,
  placeholder,
  filter = true,
  label,
  renderOption,
  numeric = false,
  onStep,
  className = '',
  listWidth,
}: {
  value: string;
  options: ComboOption[] | (() => ComboOption[]);
  onCommit: (v: string) => void;
  width?: number;
  placeholder?: string;
  filter?: boolean;
  label?: string;
  renderOption?: (o: ComboOption) => ReactNode;
  numeric?: boolean;
  onStep?: (dir: 1 | -1) => void;
  className?: string;
  listWidth?: number;
}) {
  const wrapRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const [text, setText] = useState(value);
  const [open, setOpen] = useState(false);
  const [typed, setTyped] = useState(false);
  const [active, setActive] = useState(-1);
  const focused = useRef(false);

  useEffect(() => {
    if (!focused.current) setText(value);
  }, [value]);

  const all = useMemo(() => (open ? (typeof options === 'function' ? options() : options) : []), [open, options]);
  const shown = useMemo(() => {
    if (!filter || !typed || !text) return all;
    const q = text.toLowerCase();
    const starts = all.filter((o) => (o.label ?? o.value).toLowerCase().startsWith(q));
    const contains = all.filter((o) => !(o.label ?? o.value).toLowerCase().startsWith(q) && (o.label ?? o.value).toLowerCase().includes(q));
    return [...starts, ...contains];
  }, [all, text, typed, filter]);

  useEffect(() => {
    if (!open) return;
    const i = typed ? (shown.length ? 0 : -1) : shown.findIndex((o) => o.value.toLowerCase() === value.toLowerCase());
    setActive(i);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, shown]);

  useEffect(() => {
    if (open && active >= 0) listRef.current?.querySelector(`[data-idx="${active}"]`)?.scrollIntoView({ block: 'nearest' });
  }, [active, open]);

  const commit = (v: string) => {
    setOpen(false);
    setTyped(false);
    const trimmed = v.trim();
    if (trimmed && trimmed !== value) onCommit(trimmed);
    else setText(value);
  };

  let lastGroup: string | undefined;
  return (
    <div ref={wrapRef} className={`combo${open ? ' open' : ''} ${className}`} style={{ width }}>
      <input
        ref={inputRef}
        className="combo-input"
        aria-label={label}
        value={text}
        placeholder={placeholder}
        spellCheck={false}
        onFocus={(e) => {
          focused.current = true;
          e.currentTarget.select();
        }}
        onBlur={() => {
          focused.current = false;
          setTimeout(() => {
            if (!wrapRef.current?.contains(document.activeElement)) {
              setOpen(false);
              setTyped(false);
              setText(value);
            }
          }, 120);
        }}
        onChange={(e) => {
          setText(e.target.value);
          setTyped(true);
          if (filter) setOpen(true);
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            const pick = open && active >= 0 && shown[active] ? shown[active].value : text;
            commit(pick);
            inputRef.current?.blur();
          } else if (e.key === 'Escape') {
            setOpen(false);
            setTyped(false);
            setText(value);
            inputRef.current?.blur();
          } else if (e.key === 'ArrowDown') {
            e.preventDefault();
            if (numeric && onStep && !open) onStep(-1);
            else if (!open) setOpen(true);
            else setActive((a) => Math.min(shown.length - 1, a + 1));
          } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            if (numeric && onStep && !open) onStep(1);
            else setActive((a) => Math.max(0, a - 1));
          }
        }}
      />
      <button
        type="button"
        tabIndex={-1}
        className="combo-toggle"
        aria-label="Show options"
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => {
          setTyped(false);
          setOpen((o) => !o);
          inputRef.current?.focus();
        }}
      >
        <ChevronDown size={13} />
      </button>
      <Popover open={open} anchor={wrapRef.current} onClose={() => setOpen(false)} ignore={[wrapRef.current]} placement="bottom-start" matchWidth>
        <div ref={listRef} className="combo-list thin-scroll" style={listWidth ? { width: listWidth } : undefined} onMouseDown={(e) => e.preventDefault()}>
          {shown.length === 0 && <div className="combo-empty">No matches — press Enter to use “{text}”</div>}
          {shown.map((o, i) => {
            const header = o.group && o.group !== lastGroup ? o.group : null;
            lastGroup = o.group;
            return (
              <div key={`${o.group ?? ''}:${o.value}`}>
                {header && <div className="combo-group">{header}</div>}
                <div
                  data-idx={i}
                  className={`combo-opt${i === active ? ' active' : ''}${o.value.toLowerCase() === value.toLowerCase() ? ' selected' : ''}`}
                  onMouseEnter={() => setActive(i)}
                  onClick={() => {
                    commit(o.value);
                    inputRef.current?.blur();
                  }}
                  style={o.style}
                >
                  {renderOption ? renderOption(o) : (o.label ?? o.value)}
                  {o.hint && <span className="combo-hint">{o.hint}</span>}
                </div>
              </div>
            );
          })}
        </div>
      </Popover>
    </div>
  );
}

/* ------------------------------------------------------------------ select */

export function Select<T extends string>({
  value,
  options,
  onChange,
  width,
  label,
}: {
  value: T;
  options: Array<{ value: T; label: string }>;
  onChange: (v: T) => void;
  width?: number | string;
  label?: string;
}) {
  return (
    <div className="select" style={{ width }}>
      <select aria-label={label} value={value} onChange={(e) => onChange(e.target.value as T)}>
        {options.map((o) => (
          <option key={o.value} value={o.value}>
            {o.label}
          </option>
        ))}
      </select>
      <ChevronDown size={14} className="select-chevron" />
    </div>
  );
}

export function Field({ label, children, hint, inline }: { label: ReactNode; children: ReactNode; hint?: ReactNode; inline?: boolean }) {
  return (
    <label className={`field${inline ? ' field-inline' : ''}`}>
      <span className="field-label">{label}</span>
      {children}
      {hint && <span className="field-hint">{hint}</span>}
    </label>
  );
}
