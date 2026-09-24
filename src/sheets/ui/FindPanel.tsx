import { ChevronDown, ChevronUp, X } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { Checkbox, Select } from '@/ui/controls';
import { toast } from '@/ui/toast';
import { cellKey } from '../model/address';
import type { SheetDoc } from '../doc';
import { findAll, replaceAll, type FindHit, type FindOptions } from '../ops/data';
import type { SheetUI } from './controller';

export function FindPanel({ doc, ui, mode, onClose }: { doc: SheetDoc; ui: SheetUI; mode: 'find' | 'replace'; onClose: () => void }) {
  const [q, setQ] = useState('');
  const [rep, setRep] = useState('');
  const [showReplace, setShowReplace] = useState(mode === 'replace');
  const [opts, setOpts] = useState<FindOptions>({ matchCase: false, wholeCell: false, lookIn: 'values', scope: 'sheet' });
  const [hits, setHits] = useState<FindHit[]>([]);
  const [idx, setIdx] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => setShowReplace(mode === 'replace'), [mode]);
  useEffect(() => inputRef.current?.focus(), []);

  useEffect(() => {
    const t = setTimeout(() => {
      const h = findAll(doc, q, opts);
      setHits(h);
      setIdx(0);
      ui.findHits = new Set(h.filter((x) => x.sheet === doc.sheet.id).map((x) => cellKey(x.r, x.c)));
      ui.emit();
    }, 120);
    return () => clearTimeout(t);
  }, [q, opts, doc, ui, doc.version]);

  useEffect(
    () => () => {
      ui.findHits = null;
      ui.emit();
    },
    [ui],
  );

  const go = (i: number) => {
    if (!hits.length) return;
    const n = (i + hits.length) % hits.length;
    setIdx(n);
    const h = hits[n];
    const sIdx = doc.wb.sheets.findIndex((s) => s.id === h.sheet);
    if (sIdx !== doc.wb.activeSheet) doc.setActiveSheet(sIdx);
    ui.selectCell(h.r, h.c);
  };

  return (
    <div className="find-panel sheet-find" role="search" onKeyDown={(e) => e.key === 'Escape' && onClose()}>
      <div className="find-row">
        <input
          ref={inputRef}
          className="input input-sm"
          placeholder="Find in sheet"
          value={q}
          onChange={(e) => setQ(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') go(e.shiftKey ? idx - 1 : idx + (hits.length && ui.findHits ? 1 : 0));
          }}
        />
        <span className="find-count muted">{q ? (hits.length ? `${idx + 1} of ${hits.length}` : 'No results') : ''}</span>
        <button className="icon-btn icon-btn-sm" aria-label="Previous" onClick={() => go(idx - 1)}>
          <ChevronUp size={15} />
        </button>
        <button className="icon-btn icon-btn-sm" aria-label="Next" onClick={() => go(idx + 1)}>
          <ChevronDown size={15} />
        </button>
        <button className="btn btn-sm btn-ghost" onClick={() => setShowReplace((s) => !s)}>
          Replace
        </button>
        <button className="icon-btn icon-btn-sm" aria-label="Close" onClick={onClose}>
          <X size={15} />
        </button>
      </div>
      {showReplace && (
        <div className="find-row">
          <input className="input input-sm" placeholder="Replace with" value={rep} onChange={(e) => setRep(e.target.value)} />
          <button
            className="btn btn-sm"
            disabled={!hits.length}
            onClick={() => {
              const h = hits[idx];
              if (!h) return;
              replaceAll(doc, q, rep, opts, [h]);
              go(idx);
            }}
          >
            Replace
          </button>
          <button
            className="btn btn-sm"
            disabled={!hits.length}
            onClick={() => {
              const n = replaceAll(doc, q, rep, opts);
              toast.success(`Replaced ${n} ${n === 1 ? 'cell' : 'cells'}`);
            }}
          >
            Replace all
          </button>
        </div>
      )}
      <div className="find-row find-opts">
        <Checkbox checked={!!opts.matchCase} onChange={(v) => setOpts((o) => ({ ...o, matchCase: v }))} label="Match case" />
        <Checkbox checked={!!opts.wholeCell} onChange={(v) => setOpts((o) => ({ ...o, wholeCell: v }))} label="Entire cell" />
        <Checkbox checked={!!opts.regex} onChange={(v) => setOpts((o) => ({ ...o, regex: v }))} label="Regex" />
        <Select value={opts.lookIn ?? 'values'} onChange={(v) => setOpts((o) => ({ ...o, lookIn: v }))} options={[{ value: 'values', label: 'Values' }, { value: 'formulas', label: 'Formulas' }]} width={100} />
        <Select value={opts.scope ?? 'sheet'} onChange={(v) => setOpts((o) => ({ ...o, scope: v }))} options={[{ value: 'sheet', label: 'This sheet' }, { value: 'workbook', label: 'Workbook' }, { value: 'selection', label: 'Selection' }]} width={120} />
      </div>
    </div>
  );
}
