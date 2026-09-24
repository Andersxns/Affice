import { ArrowDownAZ, ArrowUpZA, FilterX, Search } from 'lucide-react';
import { useMemo, useState } from 'react';
import { Checkbox, Select } from '@/ui/controls';
import { Popover } from '@/ui/popover';
import type { ColumnFilter } from '../model/types';
import type { SheetDoc } from '../doc';
import { filterChoices, setColumnFilter, sortRange } from '../ops/data';
import { listItems } from './controller';

type CondOp = NonNullable<ColumnFilter['condition']>['op'];

export function FilterMenu({ doc, col, anchor, onClose }: { doc: SheetDoc; col: number; anchor: DOMRect; onClose: () => void }) {
  const sheet = doc.sheet;
  const af = sheet.filter!;
  const current = af.columns[col - af.range.c1];
  const choices = useMemo(() => filterChoices(doc, sheet, col), [doc, sheet, col]);
  const [q, setQ] = useState('');
  const [picked, setPicked] = useState<Set<string>>(() => new Set(current?.values ?? choices.map((c) => c.text)));
  const [op, setOp] = useState<CondOp | 'none'>(current?.condition?.op ?? 'none');
  const [v1, setV1] = useState(current?.condition?.value ?? '');
  const [v2, setV2] = useState(current?.condition?.value2 ?? '');
  const shown = choices.filter((c) => !q || c.text.toLowerCase().includes(q.toLowerCase()));
  const allOn = shown.every((c) => picked.has(c.text));
  const numeric = choices.filter((c) => c.text).every((c) => /^[-$€£¥(]?[\d,.]+%?\)?$/.test(c.text.trim()));

  const apply = () => {
    const f: ColumnFilter = {};
    if (picked.size !== choices.length) f.values = [...picked];
    if (op !== 'none' && v1 !== '') f.condition = { op, value: v1, value2: v2 || undefined };
    setColumnFilter(doc, col, f.values || f.condition ? f : null);
    onClose();
  };

  const sort = (desc: boolean) => {
    sortRange(doc, af.range, [{ col, desc }], true);
    onClose();
  };

  return (
    <Popover open anchor={anchor} onClose={onClose} placement="bottom-start">
      <div className="filter-menu">
        <button className="menu-item" onClick={() => sort(false)}>
          <ArrowDownAZ size={15} /> {numeric ? 'Sort smallest to largest' : 'Sort A to Z'}
        </button>
        <button className="menu-item" onClick={() => sort(true)}>
          <ArrowUpZA size={15} /> {numeric ? 'Sort largest to smallest' : 'Sort Z to A'}
        </button>
        <button
          className="menu-item"
          disabled={!current}
          onClick={() => {
            setColumnFilter(doc, col, null);
            onClose();
          }}
        >
          <FilterX size={15} /> Clear filter
        </button>
        <div className="menu-sep" />
        <div className="filter-cond">
          <Select
            value={op}
            onChange={(v) => setOp(v)}
            options={[
              { value: 'none', label: numeric ? 'Number filter…' : 'Text filter…' },
              ...(numeric
                ? [
                    { value: 'equal' as const, label: 'Equals' },
                    { value: 'notEqual' as const, label: 'Does not equal' },
                    { value: 'greater' as const, label: 'Greater than' },
                    { value: 'greaterEqual' as const, label: 'Greater than or equal to' },
                    { value: 'less' as const, label: 'Less than' },
                    { value: 'lessEqual' as const, label: 'Less than or equal to' },
                    { value: 'between' as const, label: 'Between' },
                  ]
                : [
                    { value: 'contains' as const, label: 'Contains' },
                    { value: 'notContains' as const, label: 'Does not contain' },
                    { value: 'begins' as const, label: 'Begins with' },
                    { value: 'ends' as const, label: 'Ends with' },
                    { value: 'equal' as const, label: 'Equals' },
                    { value: 'notEqual' as const, label: 'Does not equal' },
                  ]),
            ]}
          />
          {op !== 'none' && <input className="input input-sm" value={v1} onChange={(e) => setV1(e.target.value)} placeholder="Value" />}
          {op === 'between' && <input className="input input-sm" value={v2} onChange={(e) => setV2(e.target.value)} placeholder="and" />}
        </div>
        <div className="filter-search">
          <Search size={14} />
          <input className="input input-sm" placeholder="Search" value={q} onChange={(e) => setQ(e.target.value)} />
        </div>
        <div className="filter-values thin-scroll">
          <Checkbox
            checked={allOn}
            onChange={(v) =>
              setPicked((p) => {
                const n = new Set(p);
                for (const c of shown) {
                  if (v) n.add(c.text);
                  else n.delete(c.text);
                }
                return n;
              })
            }
            label={<b>(Select all)</b>}
          />
          {shown.slice(0, 1000).map((c) => (
            <Checkbox
              key={c.text}
              checked={picked.has(c.text)}
              onChange={(v) =>
                setPicked((p) => {
                  const n = new Set(p);
                  if (v) n.add(c.text);
                  else n.delete(c.text);
                  return n;
                })
              }
              label={
                <span className="filter-value">
                  {c.text === '' ? <i>(Blanks)</i> : c.text}
                  <span className="muted"> {c.count}</span>
                </span>
              }
            />
          ))}
        </div>
        <div className="filter-foot">
          <button className="btn btn-sm" onClick={onClose}>
            Cancel
          </button>
          <button className="btn btn-sm btn-primary" disabled={!picked.size && op === 'none'} onClick={apply}>
            OK
          </button>
        </div>
      </div>
    </Popover>
  );
}

export function ListPicker({ doc, anchor, onPick, onClose }: { doc: SheetDoc; anchor: DOMRect; onPick: (v: string) => void; onClose: () => void }) {
  const { r, c } = doc.sel.active;
  const rule = doc.sheet.validations.find((v) => v.type === 'list' && v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
  const items = rule ? listItems(doc, rule.values, r, c) : [];
  const current = doc.displayText(doc.sheet, r, c).text;
  return (
    <Popover open anchor={anchor} onClose={onClose} placement="bottom-start">
      <div className="list-picker thin-scroll" role="listbox" style={{ minWidth: Math.max(120, anchor.width) }}>
        {items.map((it) => (
          <button key={it} role="option" aria-selected={it === current} className={`menu-item${it === current ? ' active' : ''}`} onClick={() => onPick(it)}>
            {it}
          </button>
        ))}
        {!items.length && <div className="muted small pad">The list is empty.</div>}
      </div>
    </Popover>
  );
}
